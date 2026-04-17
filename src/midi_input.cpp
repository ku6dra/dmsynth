#include "midi_input.h"
#include <cstdio>
#include <cstring>

MidiInput::MidiInput() {}

MidiInput::~MidiInput()
{
    Close();
}

std::vector<MidiDeviceInfo> MidiInput::EnumerateDevices()
{
    std::vector<MidiDeviceInfo> devices;
    UINT count = midiInGetNumDevs();
    for (UINT i = 0; i < count; i++)
    {
        MIDIINCAPSW caps = {};
        if (midiInGetDevCapsW(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR)
        {
            MidiDeviceInfo info;
            info.id = i;
            info.name = caps.szPname;
            info.mid = caps.wMid;
            info.pid = caps.wPid;
            devices.push_back(info);
        }
    }
    return devices;
}

bool MidiInput::Open(UINT deviceId)
{
    Close();

    MMRESULT result = midiInOpen(&m_hMidiIn, deviceId, reinterpret_cast<DWORD_PTR>(MidiInProc),
                                 reinterpret_cast<DWORD_PTR>(this), CALLBACK_FUNCTION);
    if (result != MMSYSERR_NOERROR)
    {
        wchar_t errMsg[256];
        midiInGetErrorTextW(result, errMsg, 256);
        fwprintf(stderr, L"midiInOpen failed: %s\n", errMsg);
        m_hMidiIn = nullptr;
        return false;
    }

    PrepareSysExBuffers();

    result = midiInStart(m_hMidiIn);
    // Capture timeGetTime() as close to midiInStart() as possible so callers
    // can convert midiInProc's dwParam2 (ms since midiInStart) back to the
    // absolute timeGetTime() domain.
    m_startTimeMs = timeGetTime();
    if (result != MMSYSERR_NOERROR)
    {
        wchar_t errMsg[256];
        midiInGetErrorTextW(result, errMsg, 256);
        fwprintf(stderr, L"midiInStart failed: %s\n", errMsg);
        UnprepareSysExBuffers();
        midiInClose(m_hMidiIn);
        m_hMidiIn = nullptr;
        return false;
    }

    return true;
}

void MidiInput::Close()
{
    if (m_hMidiIn)
    {
        m_closing = true;
        midiInStop(m_hMidiIn);
        midiInReset(m_hMidiIn);
        UnprepareSysExBuffers();
        midiInClose(m_hMidiIn);
        m_hMidiIn = nullptr;
        m_closing = false;
    }
}

void MidiInput::PrepareSysExBuffers()
{
    for (int i = 0; i < SYSEX_BUFFER_COUNT; i++)
    {
        memset(&m_sysExHeaders[i], 0, sizeof(MIDIHDR));
        m_sysExHeaders[i].lpData = reinterpret_cast<LPSTR>(m_sysExBuffers[i]);
        m_sysExHeaders[i].dwBufferLength = SYSEX_BUFFER_SIZE;
        midiInPrepareHeader(m_hMidiIn, &m_sysExHeaders[i], sizeof(MIDIHDR));
        midiInAddBuffer(m_hMidiIn, &m_sysExHeaders[i], sizeof(MIDIHDR));
    }
}

void MidiInput::UnprepareSysExBuffers()
{
    for (int i = 0; i < SYSEX_BUFFER_COUNT; i++)
    {
        if (m_sysExHeaders[i].dwFlags & MHDR_PREPARED)
        {
            midiInUnprepareHeader(m_hMidiIn, &m_sysExHeaders[i], sizeof(MIDIHDR));
        }
    }
}

void CALLBACK MidiInput::MidiInProc(HMIDIIN hMidiIn, UINT wMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1,
                                    DWORD_PTR dwParam2)
{
    auto* self = reinterpret_cast<MidiInput*>(dwInstance);

    switch (wMsg)
    {
    case MIM_DATA:
        self->HandleShortMessage(dwParam1, dwParam2);
        break;
    case MIM_LONGDATA:
        self->HandleSysEx(reinterpret_cast<MIDIHDR*>(dwParam1), dwParam2);
        break;
    case MIM_CLOSE:
        // Revert MMCSS thread priority boost (must be called from the callback thread)
        if (self->m_mmcssHandle)
        {
            AvRevertMmThreadCharacteristics(self->m_mmcssHandle);
            self->m_mmcssHandle = nullptr;
        }
        break;
    case MIM_ERROR:
        // Invalid short message received - ignore
        break;
    case MIM_LONGERROR:
        // Invalid SysEx - re-queue the buffer (unless closing)
        if (!self->m_closing)
        {
            auto* pHeader = reinterpret_cast<MIDIHDR*>(dwParam1);
            if (self->m_hMidiIn)
            {
                midiInAddBuffer(self->m_hMidiIn, pHeader, sizeof(MIDIHDR));
            }
        }
        break;
    }
}

void MidiInput::EnsureCallbackThreadBoosted()
{
    if (m_mmcssHandle)
        return;

    DWORD taskIndex = 0;
    m_mmcssHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &taskIndex);
    if (m_mmcssHandle)
    {
        // Highest priority documented for Pro Audio under MMCSS.
        AvSetMmThreadPriority(m_mmcssHandle, AVRT_PRIORITY_CRITICAL);
    }
    // Belt-and-braces: also pin the OS thread priority to TIME_CRITICAL so
    // core-parking / scheduler quirks are less likely to deschedule us mid-
    // callback. CPU use is well under 1%, so starvation risk is negligible.
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
}

void MidiInput::HandleShortMessage(DWORD_PTR dwParam1, DWORD_PTR dwParam2)
{
    EnsureCallbackThreadBoosted();

    uint8_t status = static_cast<uint8_t>(dwParam1 & 0xFF);

    // Filter System Real-Time messages (0xF8-0xFF):
    // Timing Clock, Active Sensing, etc. — not needed for synthesis or logging
    if (status >= 0xF8)
        return;

    uint8_t data1 = static_cast<uint8_t>((dwParam1 >> 8) & 0xFF);
    uint8_t data2 = static_cast<uint8_t>((dwParam1 >> 16) & 0xFF);
    DWORD timestamp = static_cast<DWORD>(dwParam2);

    if (m_callback)
    {
        m_callback(status, data1, data2, timestamp);
    }
}

void MidiInput::HandleSysEx(MIDIHDR* pHeader, DWORD_PTR dwParam2)
{
    if (m_closing)
        return;

    EnsureCallbackThreadBoosted();

    if (pHeader->dwBytesRecorded > 0 && m_sysExCallback)
    {
        const uint8_t* data = reinterpret_cast<const uint8_t*>(pHeader->lpData);
        DWORD len = pHeader->dwBytesRecorded;

        // Validate framing: well-formed SysEx starts with F0, ends with F7,
        // and fits in the prepared buffer. Malformed packets (driver bug,
        // noisy cable, truncated stream) are dropped here so downstream
        // never sees a half-message.
        bool framingOk = (len <= SYSEX_BUFFER_SIZE) && (data[0] == 0xF0) && (data[len - 1] == 0xF7);
        if (framingOk)
        {
            m_sysExCallback(data, len, static_cast<DWORD>(dwParam2));
        }
        else
        {
            fwprintf(stderr, L"[MIDI] Dropped malformed SysEx (%lu bytes, first=0x%02X last=0x%02X)\n", len,
                     len > 0 ? data[0] : 0, len > 0 ? data[len - 1] : 0);
        }
    }

    // Re-queue the buffer for more SysEx data
    if (m_hMidiIn)
    {
        midiInAddBuffer(m_hMidiIn, pHeader, sizeof(MIDIHDR));
    }
}

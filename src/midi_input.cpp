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

void MidiInput::HandleShortMessage(DWORD_PTR dwParam1, DWORD_PTR dwParam2)
{
    if (!m_mmcssHandle)
    {
        DWORD taskIndex = 0;
        m_mmcssHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &taskIndex);
    }

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

    if (!m_mmcssHandle)
    {
        DWORD taskIndex = 0;
        m_mmcssHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &taskIndex);
    }

    if (pHeader->dwBytesRecorded > 0 && m_sysExCallback)
    {
        m_sysExCallback(reinterpret_cast<const uint8_t*>(pHeader->lpData), pHeader->dwBytesRecorded,
                        static_cast<DWORD>(dwParam2));
    }

    // Re-queue the buffer for more SysEx data
    if (m_hMidiIn)
    {
        midiInAddBuffer(m_hMidiIn, pHeader, sizeof(MIDIHDR));
    }
}

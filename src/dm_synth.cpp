#include <initguid.h>
#include "dm_synth.h"
#include <dsound.h>
#include <cstring>

// XG Reverb Type (MSB, LSB) → GS Reverb Macro type
// GS: 00=Room1, 01=Room2, 02=Room3, 03=Hall1, 04=Hall2, 05=Plate, 06=Delay, 07=PanDelay
static uint8_t MapXgReverbToGs(uint8_t msb, uint8_t lsb)
{
    switch (msb)
    {
    case 0x00:
        return 0x00;                     // No Effect → Room1 (lightest reverb)
    case 0x01:                           // Hall
        return (lsb >= 1) ? 0x04 : 0x03; // Hall1 / Hall2
    case 0x02:                           // Room
        if (lsb >= 2)
            return 0x02; // Room3
        if (lsb >= 1)
            return 0x01; // Room2
        return 0x00;     // Room1
    case 0x03:
        return 0x05; // Stage → Plate
    case 0x04:
        return 0x05; // Plate → Plate
    case 0x05:
        return 0x06; // Delay LCR → Delay
    case 0x06:
        return 0x06; // Delay LR → Delay
    case 0x07:
        return 0x06; // Echo → Delay
    case 0x08:
        return 0x07; // Cross Delay → PanDelay
    default:
        return 0x04; // Unknown → Hall2 (default)
    }
}

// XG Chorus Type (MSB, LSB) → GS Chorus Macro type
// GS: 00=Chorus1, 01=Chorus2, 02=Chorus3, 03=Chorus4, 04=FB Chorus, 05=Flanger, 06=Short Delay, 07=Short Delay FB
static uint8_t MapXgChorusToGs(uint8_t msb, uint8_t lsb)
{
    switch (msb)
    {
    case 0x00:
        return 0x00; // No Effect → Chorus1 (lightest)
    case 0x41:       // Chorus
        if (lsb >= 3)
            return 0x03; // Chorus4
        if (lsb >= 2)
            return 0x02; // Chorus3
        if (lsb >= 1)
            return 0x01; // Chorus2
        return 0x00;     // Chorus1
    case 0x42:
        return 0x04; // Celeste → FB Chorus
    case 0x43:
        return 0x05; // Flanger → Flanger
    case 0x44:
        return 0x02; // Symphonic → Chorus3
    case 0x45:
        return 0x06; // Phaser → Short Delay (approximation)
    default:
        return 0x02; // Unknown → Chorus3 (default)
    }
}

DmSynth::DmSynth() {}

DmSynth::~DmSynth()
{
    Shutdown();
}

bool DmSynth::Initialize()
{
    if (m_initialized)
        return true;

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE)
    {
        fwprintf(stderr, L"CoInitializeEx failed: 0x%08X\n", hr);
        return false;
    }
    m_comInitialized = SUCCEEDED(hr);

    auto cleanup = [&]() {
        if (m_pDirectMusic)
        {
            m_pDirectMusic->Activate(FALSE);
            m_pDirectMusic->Release();
            m_pDirectMusic = nullptr;
        }
        if (m_pDirectSound)
        {
            m_pDirectSound->Release();
            m_pDirectSound = nullptr;
        }
        if (m_comInitialized)
        {
            CoUninitialize();
            m_comInitialized = false;
        }
    };

    // DirectMusic8 initialization (bypassing IDirectMusicPerformance8)
    hr = CoCreateInstance(CLSID_DirectMusic, nullptr, CLSCTX_INPROC, IID_IDirectMusic8,
                          reinterpret_cast<void**>(&m_pDirectMusic));

    if (FAILED(hr))
    {
        fwprintf(stderr, L"Failed to create IDirectMusic8: 0x%08X\n", hr);
        cleanup();
        return false;
    }

    // Explicitly create DirectSound8 and set cooperative level
    // Passing nullptr to SetDirectSound causes E_UNEXPECTED on Activate on modern Windows
    hr = DirectSoundCreate8(nullptr, &m_pDirectSound, nullptr);
    if (FAILED(hr))
    {
        fwprintf(stderr, L"DirectSoundCreate8 failed: 0x%08X\n", hr);
        cleanup();
        return false;
    }

    HWND hwnd = GetDesktopWindow();

    hr = m_pDirectSound->SetCooperativeLevel(hwnd, DSSCL_PRIORITY);
    if (FAILED(hr))
    {
        fwprintf(stderr, L"SetCooperativeLevel failed: 0x%08X\n", hr);
        cleanup();
        return false;
    }

    hr = m_pDirectMusic->SetDirectSound(m_pDirectSound, hwnd);
    if (FAILED(hr))
    {
        fwprintf(stderr, L"SetDirectSound failed: 0x%08X\n", hr);
        cleanup();
        return false;
    }

    hr = m_pDirectMusic->Activate(TRUE);
    if (FAILED(hr))
    {
        fwprintf(stderr, L"Failed to activate DirectMusic: 0x%08X\n", hr);
        cleanup();
        return false;
    }

    m_initialized = true;
    return true;
}

std::vector<SynthPortInfo> DmSynth::EnumeratePorts()
{
    std::vector<SynthPortInfo> ports;
    if (!m_initialized)
        return ports;

    DMUS_PORTCAPS portCaps;
    ZeroMemory(&portCaps, sizeof(DMUS_PORTCAPS));
    portCaps.dwSize = sizeof(DMUS_PORTCAPS);

    DWORD index = 0;
    while (m_pDirectMusic->EnumPort(index, &portCaps) == S_OK)
    {
        SynthPortInfo info;
        info.index = index;
        info.guid = portCaps.guidPort;
        info.description = portCaps.wszDescription;
        info.dwClass = portCaps.dwClass;
        info.maxVoices = portCaps.dwMaxVoices;
        info.maxChannelGroups = portCaps.dwMaxChannelGroups;
        info.flags = portCaps.dwFlags;
        info.isSoftwareSynth = (portCaps.dwFlags & DMUS_PC_SOFTWARESYNTH) != 0;

        ports.push_back(info);
        index++;
    }

    return ports;
}

bool DmSynth::CreateSynthPort(const GUID& portGuid, const SynthConfig& config)
{
    if (!m_initialized)
        return false;

    if (m_pPort)
    {
        m_pPort->Release();
        m_pPort = nullptr;
    }

    if (m_pMusicBuffer)
    {
        m_pMusicBuffer->Release();
        m_pMusicBuffer = nullptr;
    }

    m_activeConfig = config;

    // Reset bank tracking
    memset(m_bankMSB, 0, sizeof(m_bankMSB));
    memset(m_bankLSB, 0, sizeof(m_bankLSB));

    // Reset drum channel tracking (A10=index 9, B10=index 25 are drum by default)
    memset(m_drumChannel, 0, sizeof(m_drumChannel));
    m_drumChannel[9] = true;  // Group 1 (A), CH10
    m_drumChannel[25] = true; // Group 2 (B), CH10

    DMUS_PORTPARAMS8 params;
    ZeroMemory(&params, sizeof(DMUS_PORTPARAMS8));
    params.dwSize = sizeof(DMUS_PORTPARAMS8);
    params.dwValidParams = DMUS_PORTPARAMS_VOICES | DMUS_PORTPARAMS_CHANNELGROUPS | DMUS_PORTPARAMS_SAMPLERATE |
                           DMUS_PORTPARAMS_EFFECTS | DMUS_PORTPARAMS_AUDIOCHANNELS;
    params.dwVoices = config.voices;
    params.dwChannelGroups = config.channelGroups;
    params.dwSampleRate = config.sampleRate;
    params.dwAudioChannels = config.audioChannels;
    params.dwEffectFlags = DMUS_EFFECT_REVERB | DMUS_EFFECT_CHORUS | DMUS_EFFECT_DELAY;

    HRESULT hr = S_OK;
    bool success = false;
    uint32_t attemptRate = config.sampleRate;

    while (!success)
    {
        if (m_pPort)
        {
            m_pPort->Release();
            m_pPort = nullptr;
        }

        params.dwSampleRate = attemptRate;
        hr = m_pDirectMusic->CreatePort(portGuid, &params, &m_pPort, nullptr);

        if (SUCCEEDED(hr))
        {
            // CreatePort may return S_FALSE if parameters were negotiated.
            // The actual values are written back into params.
            uint32_t negotiatedRate = params.dwSampleRate;
            if (hr == S_FALSE)
            {
                fwprintf(stderr, L"Port parameters negotiated (requested: %lu Hz -> actual: %lu Hz)\n", attemptRate,
                         negotiatedRate);
            }

            hr = m_pPort->Activate(TRUE);
            if (SUCCEEDED(hr))
            {
                success = true;
                m_activeConfig.sampleRate = negotiatedRate;
            }
            else
            {
                fwprintf(stderr, L"Failed to activate port: 0x%08X (sample rate: %lu Hz)\n", hr, attemptRate);
            }
        }
        else
        {
            fwprintf(stderr, L"Failed to create port: 0x%08X (sample rate: %lu Hz)\n", hr, attemptRate);
        }

        if (!success)
        {
            if (attemptRate != 44100)
            {
                fwprintf(stderr, L"%lu Hz not supported, retrying at 44100 Hz...\n", attemptRate);
                attemptRate = 44100;
            }
            else
            {
                fwprintf(stderr, L"Failed to create synthesizer port (all sample rates failed)\n");
                return false;
            }
        }
    }

    if (!DownloadGMInstruments())
    {
        fwprintf(stderr, L"Failed to download GM instruments.\n");
    }

    // Create Music Buffer for Playback
    DMUS_BUFFERDESC bufferDesc;
    ZeroMemory(&bufferDesc, sizeof(DMUS_BUFFERDESC));
    bufferDesc.dwSize = sizeof(DMUS_BUFFERDESC);
    bufferDesc.dwFlags = 0;
    bufferDesc.guidBufferFormat = GUID_NULL;
    bufferDesc.cbBuffer = 4096;

    hr = m_pDirectMusic->CreateMusicBuffer(&bufferDesc, &m_pMusicBuffer, nullptr);
    if (FAILED(hr))
    {
        fwprintf(stderr, L"Failed to create MusicBuffer: 0x%08X\n", hr);
        return false;
    }

    {
        std::lock_guard<std::mutex> lock(m_mutex);
        SendDefaultEffects();
    }
    return true;
}

bool DmSynth::DownloadGMInstruments()
{
    if (!m_pPort)
        return false;

    if (!m_pLoader)
    {
        HRESULT hr = CoCreateInstance(CLSID_DirectMusicLoader, nullptr, CLSCTX_INPROC, IID_IDirectMusicLoader8,
                                      reinterpret_cast<void**>(&m_pLoader));
        if (FAILED(hr))
        {
            fwprintf(stderr, L"Failed to create DirectMusicLoader: 0x%08X\n", hr);
            return false;
        }
    }

    if (!m_pDefaultDls)
    {
        wchar_t resolvedPath[MAX_PATH];
        if (m_activeConfig.dlsPath.empty())
        {
            // Default: system gm.dls
            GetSystemDirectoryW(resolvedPath, MAX_PATH);
            wcscat_s(resolvedPath, L"\\drivers\\gm.dls");
        }
        else
        {
            // User-specified DLS file — resolve to absolute path
            if (!GetFullPathNameW(m_activeConfig.dlsPath.c_str(), MAX_PATH, resolvedPath, nullptr))
            {
                fwprintf(stderr, L"Failed to resolve DLS path: %s\n", m_activeConfig.dlsPath.c_str());
                return false;
            }
        }

        HRESULT hr = m_pLoader->LoadObjectFromFile(CLSID_DirectMusicCollection, IID_IDirectMusicCollection8,
                                                   resolvedPath, reinterpret_cast<void**>(&m_pDefaultDls));
        if (FAILED(hr))
        {
            fwprintf(stderr, L"Failed to load DLS file: %s (0x%08X)\n", resolvedPath, hr);
            return false;
        }
        fwprintf(stdout, L"DLS: %s\n", resolvedPath);
    }

    m_availablePatches.clear();
    m_downloadedInstruments.clear();

    DMUS_NOTERANGE noteRange;
    noteRange.dwLowNote = 0;
    noteRange.dwHighNote = 127;

    // Enumerate all instruments in the collection
    DWORD index = 0;
    DWORD patch = 0;
    WCHAR name[64];
    while (m_pDefaultDls->EnumInstrument(index, &patch, name, 64) == S_OK)
    {
        IDirectMusicInstrument* pInst = nullptr;
        if (SUCCEEDED(m_pDefaultDls->GetInstrument(patch, &pInst)) && pInst)
        {
            IDirectMusicDownloadedInstrument* pDownloaded = nullptr;
            if (SUCCEEDED(m_pPort->DownloadInstrument(pInst, reinterpret_cast<void**>(&pDownloaded), &noteRange, 1)) &&
                pDownloaded)
            {
                m_downloadedInstruments.push_back(pDownloaded);
                m_availablePatches.insert(patch);

                uint8_t msb = (patch >> 16) & 0xFF;
                uint8_t lsb = (patch >> 8) & 0xFF;
                uint8_t prog = patch & 0xFF;
                bool isDrum = (patch & 0x80000000) != 0;

                fwprintf(stdout, L"Loaded Instrument: %s (Bank MSB:%u LSB:%u Patch:%u%s)\n", name, msb, lsb, prog,
                         isDrum ? L" Drum" : L"");
            }
            pInst->Release();
        }
        index++;
    }

    fwprintf(stdout, L"Total instruments loaded: %zu\n", m_availablePatches.size());

    m_pPort->Compact();

    return true;
}

bool DmSynth::SendMidiMessage(uint8_t status, uint8_t data1, uint8_t data2, uint32_t channelGroup)
{
    if (!m_pPort || !m_pMusicBuffer)
        return false;
    std::lock_guard<std::mutex> lock(m_mutex);

    uint8_t msgType = status & 0xF0;
    uint8_t channel = status & 0x0F;
    uint8_t idx = channel + static_cast<uint8_t>((channelGroup - 1) * 16); // 0-31
    if (idx >= 32)
        return false;
    bool isDrum = m_drumChannel[idx];

    // Track state for bank selection and implement fallback logic
    if (msgType == 0xB0)
    { // Control Change
        if (data1 == 0x00)
        { // Bank Select MSB
            m_bankMSB[idx] = data2;
        }
        else if (data1 == 0x20)
        { // Bank Select LSB
            m_bankLSB[idx] = data2;
        }
    }
    else if (msgType == 0xC0)
    { // Program Change
        uint32_t program = data1;
        uint32_t bankMSB = m_bankMSB[idx];
        uint32_t bankLSB = m_bankLSB[idx];
        uint32_t drumFlag = isDrum ? 0x80000000 : 0;

        uint32_t requestedPatch = drumFlag | (bankMSB << 16) | (bankLSB << 8) | program;
        uint32_t bank0Patch = drumFlag | program;

        uint32_t finalPatch = 0;
        int fallbackLevel = 0;

        auto exists = [this](uint32_t p) { return m_availablePatches.count(p) > 0; };

        // Search any bank for a given program number (same drum flag)
        auto findAnyBank = [this, drumFlag](uint32_t prog) -> uint32_t {
            for (uint32_t p : m_availablePatches)
                if ((p & 0xFF) == prog && (p & 0x80000000) == drumFlag)
                    return p;
            return 0xFFFFFFFF;
        };

        if (exists(requestedPatch))
        {
            finalPatch = requestedPatch;
            fallbackLevel = 0;
        }
        else if (exists(bank0Patch))
        {
            // L1: Bank 0/0, same program
            finalPatch = bank0Patch;
            fallbackLevel = 1;
        }
        else
        {
            uint32_t anyBankPatch = findAnyBank(program);
            if (anyBankPatch != 0xFFFFFFFF)
            {
                // L2: any bank, same program
                finalPatch = anyBankPatch;
                fallbackLevel = 2;
            }
            else
            {
                uint32_t anyDefaultPatch = findAnyBank(0);
                if (anyDefaultPatch != 0xFFFFFFFF)
                {
                    // L3: any bank, program 0
                    finalPatch = anyDefaultPatch;
                    fallbackLevel = 3;
                }
                else if (!m_availablePatches.empty())
                {
                    // L4: absolute fallback, first available instrument
                    finalPatch = *m_availablePatches.begin();
                    fallbackLevel = 4;
                }
            }
        }

        if (fallbackLevel > 0)
        {
            uint32_t fMSB = (finalPatch >> 16) & 0x7F;
            uint32_t fLSB = (finalPatch >> 8) & 0xFF;
            uint32_t fProg = finalPatch & 0xFF;

            fwprintf(stdout,
                     L"[Translate] Program Change ch%u: Bank %u/%u Prog %u → Bank %u/%u Prog %u (fallback L%d)\n",
                     channel + 1, bankMSB, bankLSB, program, fMSB, fLSB, fProg, fallbackLevel);

            // Re-route to fallback:
            // 1. Send Bank Select
            m_pMusicBuffer->PackStructured(0, channelGroup, (0xB0 | channel) | (0x00 << 8) | (fMSB << 16));
            m_pMusicBuffer->PackStructured(0, channelGroup, (0xB0 | channel) | (0x20 << 8) | (fLSB << 16));

            // 2. Send Program Change
            m_pMusicBuffer->PackStructured(0, channelGroup, (0xC0 | channel) | (fProg << 8));

            m_pPort->PlayBuffer(m_pMusicBuffer);
            m_pMusicBuffer->Flush();

            // Update internal state
            m_bankMSB[idx] = (uint8_t)fMSB;
            m_bankLSB[idx] = (uint8_t)fLSB;

            return true;
        }
    }

    // Normal message processing

    // DRUM FIX: Ignore Note Off for drum channel (Channel 10)
    // Some sequences have extremely short Note On/Off durations for drums,
    // which can cause the synth to kill the voice before it even starts.
    if (isDrum)
    {
        if (msgType == 0x80 || (msgType == 0x90 && data2 == 0))
        {
            return true;
        }
    }

    DWORD midiMessage = status | (data1 << 8) | (data2 << 16);

    HRESULT hr = m_pMusicBuffer->PackStructured(0, channelGroup, midiMessage);
    if (FAILED(hr))
        return false;

    hr = m_pPort->PlayBuffer(m_pMusicBuffer);
    if (FAILED(hr))
        return false;

    m_pMusicBuffer->Flush(); // reset buffer for next message
    return true;
}

bool DmSynth::SendSysEx(const uint8_t* data, uint32_t length, uint32_t channelGroup)
{
    if (!m_pPort || !m_pMusicBuffer)
        return false;
    std::lock_guard<std::mutex> lock(m_mutex);

    // GM System On:  F0 7E 7F 09 01 F7
    // GM System Off: F0 7E 7F 09 02 F7
    // GS Reset:      F0 41 10 42 12 40 00 7F 00 41 F7
    // XG System On:  F0 43 10 4C 00 00 7E 00 F7
    // GM1: data[4]==0x01, GM2: data[4]==0x03, GM Off: data[4]==0x02
    bool isGmOff =
        (length == 6 && data[0] == 0xF0 && data[1] == 0x7E && data[3] == 0x09 && data[4] == 0x02 && data[5] == 0xF7);
    bool isGmOn = (length == 6 && data[0] == 0xF0 && data[1] == 0x7E && data[3] == 0x09 &&
                   (data[4] == 0x01 || data[4] == 0x03) && data[5] == 0xF7);
    bool isGsReset = (length == 11 && data[0] == 0xF0 && data[1] == 0x41 && data[3] == 0x42 && data[4] == 0x12 &&
                      data[5] == 0x40 && data[6] == 0x00 && data[7] == 0x7F && data[8] == 0x00 && data[10] == 0xF7);
    bool isXgOn = (length == 9 && data[0] == 0xF0 && data[1] == 0x43 && data[3] == 0x4C && data[4] == 0x00 &&
                   data[5] == 0x00 && data[6] == 0x7E && data[7] == 0x00 && data[8] == 0xF7);

    if (isGmOn || isGmOff || isGsReset || isXgOn)
    {
        // Proxy GM System On/Off and XG System On to GS Reset, which the MS GS Wavetable Synth handles natively.
        if (isGmOn)
            fwprintf(stdout, L"[Translate] GM System On → GS Reset\n");
        else if (isGmOff)
            fwprintf(stdout, L"[Translate] GM System Off → GS Reset\n");
        else if (isXgOn)
            fwprintf(stdout, L"[Translate] XG System On → GS Reset\n");
        static const uint8_t gsResetMsg[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x00, 0x7F, 0x00, 0x41, 0xF7};
        m_pMusicBuffer->PackUnstructured(0, channelGroup, sizeof(gsResetMsg), (LPBYTE)gsResetMsg);
        m_pPort->PlayBuffer(m_pMusicBuffer);
        m_pMusicBuffer->Flush();

        memset(m_bankMSB, 0, sizeof(m_bankMSB));
        memset(m_bankLSB, 0, sizeof(m_bankLSB));
        memset(m_drumChannel, 0, sizeof(m_drumChannel));
        m_drumChannel[9] = true;  // Group 1 (A), CH10
        m_drumChannel[25] = true; // Group 2 (B), CH10
        SendDefaultEffects();
        return true;
    }

    // GS USE FOR RHYTHM PART: F0 41 10 42 12 40 <CH> 15 <mode> <checksum> F7
    // <CH>: A01-A09=0x11-0x19, A10=0x10, A11-A16=0x1A-0x1F
    // <mode>: 00=Normal, 01=Drum1, 02=Drum2
    if (length == 11 && data[0] == 0xF0 && data[1] == 0x41 && data[3] == 0x42 && data[4] == 0x12 && data[5] == 0x40 &&
        data[7] == 0x15 && data[10] == 0xF7)
    {
        uint8_t chAddr = data[6];
        uint8_t mode = data[8];
        // Convert GS channel address to MIDI channel (0-based)
        // 0x10=ch10(9), 0x11-0x19=ch1-9(0-8), 0x1A-0x1F=ch11-16(10-15)
        int midiChannel = -1;
        if (chAddr == 0x10)
        {
            midiChannel = 9;
        }
        else if (chAddr >= 0x11 && chAddr <= 0x19)
        {
            midiChannel = chAddr - 0x11; // 0x11->0, 0x19->8
        }
        else if (chAddr >= 0x1A && chAddr <= 0x1F)
        {
            midiChannel = chAddr - 0x1A + 10; // 0x1A->10, 0x1F->15
        }
        if (midiChannel >= 0 && midiChannel < 16)
        {
            m_drumChannel[midiChannel + (channelGroup - 1) * 16] = (mode != 0x00);
        }
    }

    // XG Effect Parameter Change → GS Macro conversion
    // Format: F0 43 1n 4C 02 01 <addr_low> <data...> F7
    if (length >= 9 && data[0] == 0xF0 && data[1] == 0x43 && (data[2] & 0xF0) == 0x10 && data[3] == 0x4C &&
        data[4] == 0x02 && data[5] == 0x01)
    {
        uint8_t addrLow = data[6];
        if (addrLow == 0x00)
        { // Reverb Type MSB (+ optional LSB)
            uint8_t xgMsb = data[7];
            uint8_t xgLsb = (length >= 10) ? data[8] : 0;
            uint8_t gsMacro = MapXgReverbToGs(xgMsb, xgLsb);
            fwprintf(stdout, L"[Translate] XG Reverb %02X/%02X → GS Reverb macro %02X\n", xgMsb, xgLsb, gsMacro);
            SendGsReverbMacro(gsMacro);
            return true;
        }
        if (addrLow == 0x20)
        { // Chorus Type MSB (+ optional LSB)
            uint8_t xgMsb = data[7];
            uint8_t xgLsb = (length >= 10) ? data[8] : 0;
            uint8_t gsMacro = MapXgChorusToGs(xgMsb, xgLsb);
            fwprintf(stdout, L"[Translate] XG Chorus %02X/%02X → GS Chorus macro %02X\n", xgMsb, xgLsb, gsMacro);
            SendGsChorusMacro(gsMacro);
            return true;
        }
    }

    HRESULT hr = m_pMusicBuffer->PackUnstructured(0, channelGroup, length, (LPBYTE)data);
    if (FAILED(hr))
        return false;

    hr = m_pPort->PlayBuffer(m_pMusicBuffer);
    if (FAILED(hr))
        return false;

    m_pMusicBuffer->Flush();
    return true;
}

void DmSynth::ResetAllState()
{
    if (!m_pPort || !m_pMusicBuffer)
        return;
    std::lock_guard<std::mutex> lock(m_mutex);

    for (uint32_t group = 1; group <= 2; group++)
    {
        for (uint8_t ch = 0; ch < 16; ch++)
        {
            // All Sound Off (CC 120) - immediately silence all voices
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (120 << 8) | (0 << 16));
            // Reset All Controllers (CC 121)
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (121 << 8) | (0 << 16));
            // All Notes Off (CC 123)
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (123 << 8) | (0 << 16));
            // Bank Select MSB reset (CC 0)
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (0 << 8) | (0 << 16));
            // Bank Select LSB reset (CC 32)
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (32 << 8) | (0 << 16));
            // Program Change to 0
            m_pMusicBuffer->PackStructured(0, group, (0xC0 | ch) | (0 << 8));
            // Pitch Bend center
            m_pMusicBuffer->PackStructured(0, group, (0xE0 | ch) | (0 << 8) | (64 << 16));
            // GM default CC values
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (7 << 8) | (100 << 16));  // Volume = 100
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (10 << 8) | (64 << 16));  // Pan = 64 (center)
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (11 << 8) | (127 << 16)); // Expression = 127
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (91 << 8) | (40 << 16));  // Reverb Send = 40
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (93 << 8) | (0 << 16));   // Chorus Send = 0
            m_pPort->PlayBuffer(m_pMusicBuffer);
            m_pMusicBuffer->Flush();
            // Pitch Bend Sensitivity = ±2 semitones (RPN 0,0)
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (101 << 8) | (0 << 16));   // RPN MSB = 0
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (100 << 8) | (0 << 16));   // RPN LSB = 0
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (6 << 8) | (2 << 16));     // Data Entry = 2
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (38 << 8) | (0 << 16));    // Data Entry LSB = 0
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (101 << 8) | (127 << 16)); // RPN Null
            m_pMusicBuffer->PackStructured(0, group, (0xB0 | ch) | (100 << 8) | (127 << 16)); // RPN Null
            m_pPort->PlayBuffer(m_pMusicBuffer);
            m_pMusicBuffer->Flush();
        }
    }

    // Reset internal tracking state
    memset(m_bankMSB, 0, sizeof(m_bankMSB));
    memset(m_bankLSB, 0, sizeof(m_bankLSB));
    memset(m_drumChannel, 0, sizeof(m_drumChannel));
    m_drumChannel[9] = true;  // Group 1, CH10
    m_drumChannel[25] = true; // Group 2, CH10

    SendDefaultEffects();
}

// Send GS Reverb Macro SysEx: F0 41 10 42 12 40 01 30 <type> <checksum> F7
// Caller must hold m_mutex.
void DmSynth::SendGsReverbMacro(uint8_t type)
{
    uint8_t sum = 0x40 + 0x01 + 0x30 + type;
    uint8_t checksum = (128 - (sum & 0x7F)) & 0x7F;
    uint8_t sysex[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x01, 0x30, type, checksum, 0xF7};
    m_pMusicBuffer->PackUnstructured(0, 1, sizeof(sysex), sysex);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();
}

// Send GS Chorus Macro SysEx: F0 41 10 42 12 40 01 38 <type> <checksum> F7
// Caller must hold m_mutex.
void DmSynth::SendGsChorusMacro(uint8_t type)
{
    uint8_t sum = 0x40 + 0x01 + 0x38 + type;
    uint8_t checksum = (128 - (sum & 0x7F)) & 0x7F;
    uint8_t sysex[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x01, 0x38, type, checksum, 0xF7};
    m_pMusicBuffer->PackUnstructured(0, 1, sizeof(sysex), sysex);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();
}

// Send GS Delay Macro SysEx: F0 41 10 42 12 40 01 50 <type> <checksum> F7
// type: 00=Delay1, 01=Delay2, 02=Delay3, 03=Delay4,
//       04=Pan Delay1, 05=Pan Delay2, 06=Pan Delay3, 07=Pan Delay4,
//       08=Delay to Reverb, 09=Pan Repeat
// Caller must hold m_mutex.
void DmSynth::SendGsDelayMacro(uint8_t type)
{
    uint8_t sum = 0x40 + 0x01 + 0x50 + type;
    uint8_t checksum = (128 - (sum & 0x7F)) & 0x7F;
    uint8_t sysex[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x01, 0x50, type, checksum, 0xF7};
    m_pMusicBuffer->PackUnstructured(0, 1, sizeof(sysex), sysex);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();
}

// Set default GS effect types. Caller must hold m_mutex.
void DmSynth::SendDefaultEffects()
{
    // GM Master Volume = 100 (Universal Real Time SysEx: F0 7F 7F 04 01 <LSB> <MSB> F7)
    static const uint8_t masterVolMsg[] = {0xF0, 0x7F, 0x7F, 0x04, 0x01, 0x00, 0x64, 0xF7};
    m_pMusicBuffer->PackUnstructured(0, 1, sizeof(masterVolMsg), (LPBYTE)masterVolMsg);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();

    SendGsReverbMacro(0x04); // Hall2 (GS default)
    SendGsChorusMacro(0x02); // Chorus3 (GS default)
    SendGsDelayMacro(0x00);  // Delay1 (GS default)
}

void DmSynth::Shutdown()
{
    if (!m_initialized)
        return;

    if (m_pPort)
    {
        for (auto pDownloaded : m_downloadedInstruments)
        {
            if (pDownloaded)
            {
                m_pPort->UnloadInstrument(pDownloaded);
                pDownloaded->Release();
            }
        }
    }
    m_downloadedInstruments.clear();

    if (m_pDefaultDls)
    {
        m_pDefaultDls->Release();
        m_pDefaultDls = nullptr;
    }

    if (m_pLoader)
    {
        m_pLoader->Release();
        m_pLoader = nullptr;
    }

    if (m_pMusicBuffer)
    {
        m_pMusicBuffer->Release();
        m_pMusicBuffer = nullptr;
    }

    if (m_pPort)
    {
        m_pPort->Activate(FALSE);
        m_pPort->Release();
        m_pPort = nullptr;
    }

    if (m_pDirectMusic)
    {
        m_pDirectMusic->Release();
        m_pDirectMusic = nullptr;
    }

    if (m_pDirectSound)
    {
        m_pDirectSound->Release();
        m_pDirectSound = nullptr;
    }

    if (m_comInitialized)
    {
        CoUninitialize();
        m_comInitialized = false;
    }

    m_initialized = false;
}

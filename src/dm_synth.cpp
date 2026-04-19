#include <initguid.h>
#include "dm_synth.h"
#include <dsound.h>
#include <mmsystem.h>
#include <algorithm>
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

    if (m_pLatencyClock)
    {
        m_pLatencyClock->Release();
        m_pLatencyClock = nullptr;
    }
    m_anchorInitialized = false;

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

    // Acquire the port's latency clock so we can stamp MIDI buffers with a
    // reference-time that tracks actual arrival order rather than 0.
    if (FAILED(m_pPort->GetLatencyClock(&m_pLatencyClock)))
    {
        m_pLatencyClock = nullptr;
    }

    // Timing mode: scheduled by default (uses the port's latency clock to
    // absorb callback jitter). Passing --immediate forces rt=0 on every
    // outgoing message, which is required when running alongside software
    // like DirectMusic Producer — such software can throttle our
    // DirectSound buffer in a way that drags the latency clock off
    // wall-clock rate, making scheduled playback run slow. We do not
    // auto-detect this: the ratio observed in practice is noisy enough
    // that a reliable threshold is hard to pick.
    m_immediateMode = config.forceImmediateMode || !m_pLatencyClock;
    if (m_immediateMode)
    {
        fwprintf(stdout, L"Timing mode: immediate (%s)\n",
                 config.forceImmediateMode ? L"forced by --immediate" : L"no latency clock available");
    }

    {
        std::lock_guard<std::mutex> lock(m_mutex);
        SendDefaultEffects();
    }
    return true;
}

REFERENCE_TIME DmSynth::MidiMsToRefTime(DWORD absWinmmMs)
{
    if (!m_pLatencyClock || m_immediateMode)
        return 0;

    // Serialize anchor state against concurrent MIDI-in threads (e.g. port 2).
    std::lock_guard<std::mutex> lock(m_mutex);

    // Anchor once at the first message and hold it for the lifetime of the
    // port. Re-anchoring on silence was causing audible ~15 ms cbDelay
    // shifts per re-anchor because every fresh sample baked in a new ε
    // (the time between the timeGetTime and latency-clock reads, plus any
    // preemption). A stationary anchor converts that ε into a constant
    // per-session offset instead, which is inaudible. DWORD wrap of
    // timeGetTime (~49.7 days) is handled by the signed delta below.
    if (!m_anchorInitialized)
    {
        // Sandwich the latency-clock read between two timeGetTime() calls
        // and keep the sample with the smallest span. The latency clock was
        // read at some unknown point between w1 and w2, so midpoint is the
        // best estimate; the remaining error is bounded by (w2 - w1) / 2.
        //
        // Without this, a preempted GetTime() call (seen as large as ~16 ms
        // under load) gets baked into the anchor as a constant bias on every
        // subsequent lead computation — that is what was changing between
        // re-anchors.
        const int kAttempts = 8;
        DWORD bestSpan = 0xFFFFFFFFu;
        DWORD bestWinmm = 0;
        REFERENCE_TIME bestRt = 0;
        bool got = false;
        for (int i = 0; i < kAttempts; i++)
        {
            DWORD w1 = timeGetTime();
            REFERENCE_TIME rt = 0;
            if (FAILED(m_pLatencyClock->GetTime(&rt)))
                continue;
            DWORD w2 = timeGetTime();
            DWORD span = w2 - w1;
            if (!got || span < bestSpan)
            {
                bestSpan = span;
                bestWinmm = w1 + span / 2;
                bestRt = rt;
                got = true;
                if (span == 0)
                    break; // cannot do better at 1 ms resolution
            }
        }
        if (!got)
            return 0;

        m_winmmAnchorMs = bestWinmm;
        m_refAnchor = bestRt;
        m_anchorInitialized = true;
    }

    // Signed delta handles DWORD wrap (~49.7 days) correctly.
    int32_t deltaMs = static_cast<int32_t>(absWinmmMs - m_winmmAnchorMs);
    return m_refAnchor + static_cast<REFERENCE_TIME>(deltaMs) * 10000 + m_schedulingLead;
}

HRESULT DmSynth::PackStructuredWithRetry(REFERENCE_TIME rt, uint32_t group, DWORD msg)
{
    HRESULT hr = m_pMusicBuffer->PackStructured(rt, group, msg);
    if (FAILED(hr))
    {
        // Buffer likely full — drain and retry once.
        m_pPort->PlayBuffer(m_pMusicBuffer);
        m_pMusicBuffer->Flush();
        hr = m_pMusicBuffer->PackStructured(rt, group, msg);
        m_tsStats.retryCount++;
    }
    return hr;
}

HRESULT DmSynth::PackUnstructuredWithRetry(REFERENCE_TIME rt, uint32_t group, uint32_t length, const uint8_t* data)
{
    HRESULT hr = m_pMusicBuffer->PackUnstructured(rt, group, length, (LPBYTE)data);
    if (FAILED(hr))
    {
        m_pPort->PlayBuffer(m_pMusicBuffer);
        m_pMusicBuffer->Flush();
        hr = m_pMusicBuffer->PackUnstructured(rt, group, length, (LPBYTE)data);
        m_tsStats.retryCount++;
    }
    return hr;
}

void DmSynth::AccumulateTimestampStats(REFERENCE_TIME scheduledRt, bool isSysEx)
{
    if (!m_pLatencyClock || m_immediateMode)
        return;

    REFERENCE_TIME now = 0;
    if (FAILED(m_pLatencyClock->GetTime(&now)))
        return;

    // Lead = scheduled time − current latency-clock time. Negative = late
    // (DirectMusic treats as "play now"). Callback delay is the inverse
    // side of the same coin: cbDelay = schedulingLead − lead, i.e. how long
    // after MIDI arrival the message was actually packed. It's the main
    // thing to look at when deciding whether schedulingLead is big enough.
    REFERENCE_TIME lead = scheduledRt - now;
    REFERENCE_TIME cbDelay = m_schedulingLead - lead;

    if (!m_tsStats.haveSample)
    {
        m_tsStats.windowStart = now;
        m_tsStats.leadMin = lead;
        m_tsStats.leadMax = lead;
        m_tsStats.cbDelayMin = cbDelay;
        m_tsStats.cbDelayMax = cbDelay;
        m_tsStats.haveSample = true;
    }
    else
    {
        m_tsStats.leadMin = (std::min)(m_tsStats.leadMin, lead);
        m_tsStats.leadMax = (std::max)(m_tsStats.leadMax, lead);
        m_tsStats.cbDelayMin = (std::min)(m_tsStats.cbDelayMin, cbDelay);
        m_tsStats.cbDelayMax = (std::max)(m_tsStats.cbDelayMax, cbDelay);
    }
    m_tsStats.leadSum += lead;
    m_tsStats.cbDelaySum += cbDelay;
    if (isSysEx)
        m_tsStats.sysexCount++;
    else
        m_tsStats.midiCount++;
    if (lead < 0)
        m_tsStats.lateCount++;

    // Flush once per ~1 s (10,000,000 × 100ns).
    if (now - m_tsStats.windowStart >= 10000000LL)
    {
        uint32_t total = m_tsStats.midiCount + m_tsStats.sysexCount;

        // Carry over if sample size is too small to trust cbDelayMax. Advance
        // the eval anchor by ~1 s but keep accumulated stats so the next check
        // sees a longer observation window with more samples.
        if (total < kAutoTuneMinSamples)
        {
            m_tsStats.windowStart += 10000000LL;
            return;
        }

        double avgLead = static_cast<double>(m_tsStats.leadSum) / total / 10000.0;
        double avgCb = static_cast<double>(m_tsStats.cbDelaySum) / total / 10000.0;
        double cbMaxMs = static_cast<double>(m_tsStats.cbDelayMax) / 10000.0;

        // ---- Auto-tune ----
        // Tracks the steady-state callback-delay distribution; NOT a hitch
        // defense (tiny lead bumps can't absorb 10-100 ms stalls). Uses a
        // signed hysteresis accumulator: a window's signal (+1 want-up, -1
        // want-down, 0 in-band) is summed; opposing signals reset to ±1, 0s
        // don't reset. Action fires when |pending| reaches kAutoTuneHystVotes.
        double leadMsPrecise = static_cast<double>(m_schedulingLead) / 10000.0;
        double headroom = leadMsPrecise - cbMaxMs;
        bool hadLate = (m_tsStats.lateCount > 0);
        int signal = 0;
        if (hadLate || headroom < kAutoTuneTargetLowMs)
            signal = +1;
        else if (headroom > kAutoTuneTargetHighMs)
            signal = -1;

        int adjust = 0;
        if (m_autoTuneFrozen)
        {
            if (signal != 0)
            {
                m_autoTuneFrozen = false;
                m_autoTuneStableSeconds = 0;
                m_autoTunePending = signal; // seed count at 1 in that direction
            }
        }
        else
        {
            if (signal > 0)
                m_autoTunePending = (m_autoTunePending < 0) ? 1 : m_autoTunePending + 1;
            else if (signal < 0)
                m_autoTunePending = (m_autoTunePending > 0) ? -1 : m_autoTunePending - 1;
            // signal == 0: leave m_autoTunePending unchanged (= doesn't reset)

            if (m_autoTunePending >= kAutoTuneHystVotes)
            {
                adjust = +1;
                m_autoTunePending = 0;
                m_autoTuneStableSeconds = 0;
            }
            else if (m_autoTunePending <= -kAutoTuneHystVotes)
            {
                adjust = -1;
                m_autoTunePending = 0;
                m_autoTuneStableSeconds = 0;
            }
            else if (signal == 0)
            {
                // Calm in-band window: progress toward freeze even if a stale
                // unresolved vote is pending — otherwise a single lingering
                // spike could block freeze indefinitely.
                m_autoTuneStableSeconds++;
                if (m_autoTuneStableSeconds >= kAutoTuneFreezeSeconds)
                {
                    m_autoTuneFrozen = true;
                    m_autoTunePending = 0;
                }
            }
        }

        if (adjust > 0)
        {
            uint32_t shift = (std::min)(m_autoTuneConsecUp, kAutoTuneUpShiftMax);
            m_schedulingLead += kAutoTuneStepRt << shift;
            m_autoTuneConsecUp++;
        }
        else if (adjust < 0)
        {
            m_autoTuneConsecUp = 0;
            if (m_schedulingLead - kAutoTuneStepRt >= kAutoTuneMinLeadRt)
                m_schedulingLead -= kAutoTuneStepRt;
        }

#if DMSYNTH_DEBUG_TIMESTAMPS
        {
            std::lock_guard<std::mutex> slock(m_outLinesMutex);
            swprintf_s(m_statsLineBuf, _countof(m_statsLineBuf),
                       L"[TS 1s] msgs=%u sysex=%u late=%u retries=%u lead=%.2fms%s pend=%+d "
                       L"cbDelay(min/avg/max)=%.2f/%.2f/%.2f ms  "
                       L"actualLead(min/avg/max)=%.2f/%.2f/%.2f ms",
                       m_tsStats.midiCount, m_tsStats.sysexCount, m_tsStats.lateCount, m_tsStats.retryCount,
                       static_cast<double>(m_schedulingLead) / 10000.0,
                       m_autoTuneFrozen ? L"(F)" : (adjust > 0 ? L"(+)" : (adjust < 0 ? L"(-)" : L"(=)")),
                       m_autoTunePending, static_cast<double>(m_tsStats.cbDelayMin) / 10000.0, avgCb,
                       static_cast<double>(m_tsStats.cbDelayMax) / 10000.0,
                       static_cast<double>(m_tsStats.leadMin) / 10000.0, avgLead,
                       static_cast<double>(m_tsStats.leadMax) / 10000.0);
            m_statsLineReady = true;
        }
#endif
        m_tsStats = TimestampStats{};
    }
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

bool DmSynth::SendMidiMessage(uint8_t status, uint8_t data1, uint8_t data2, uint32_t channelGroup,
                              REFERENCE_TIME timestamp)
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
        else if (data1 == 0x79)
        { // Reset All Controllers: clear bank tracking for this channel
            // MIDI spec does not require bank select to reset on CC 0x79,
            // but we reset it to prevent stale bank state from affecting
            // subsequent program change fallback logic.
            m_bankMSB[idx] = 0;
            m_bankLSB[idx] = 0;
        }
    }
    else if (msgType == 0xC0)
    { // Program Change
        uint32_t program = data1;
        uint32_t bankMSB = m_bankMSB[idx];
        uint32_t bankLSB = m_bankLSB[idx];
        uint32_t drumFlag = isDrum ? 0x80000000 : 0;

        uint32_t requestedPatch = drumFlag | (bankMSB << 16) | (bankLSB << 8) | program;
        // GS tone-map fallback. LSB selects the sound-module map
        // (SC-55 / SC-88Pro / SC-8850 / ...) rather than the tone itself, so
        // it is dropped during lookup.
        //
        // For melodic channels the capitals/sub-capitals sit at MSB 0, 8,
        // 16, ... with program selecting the tone within the group, so MSB
        // rounds down to the nearest multiple of 8 (sub-capital) then to 0
        // (capital). For drum channels the kit is selected by program, and
        // SC-55's kits group the same way at PC 0 (STANDARD) / 8 (ROOM) /
        // 16 (POWER) / 24 (ELECTRONIC) / 32 (JAZZ) / 40 (BRUSH) / 48
        // (ORCHESTRA) / 56 (SFX) — so drums round PC instead of MSB.
        uint32_t subCapitalPatch, capitalPatch;
        if (isDrum)
        {
            uint32_t subCapitalPC = (program / 8) * 8;
            subCapitalPatch = drumFlag | subCapitalPC;
            capitalPatch = drumFlag; // PC 0 = STANDARD kit
        }
        else
        {
            uint32_t subCapitalMSB = (bankMSB / 8) * 8;
            subCapitalPatch = drumFlag | (subCapitalMSB << 16) | program;
            capitalPatch = drumFlag | program;
        }

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
        else if (subCapitalPatch != requestedPatch && exists(subCapitalPatch))
        {
            // L1: GS sub-capital (MSB rounded down to multiple of 8, LSB ignored)
            finalPatch = subCapitalPatch;
            fallbackLevel = 1;
        }
        else if (capitalPatch != requestedPatch && capitalPatch != subCapitalPatch && exists(capitalPatch))
        {
            // L2: GS capital (MSB 0, LSB ignored)
            finalPatch = capitalPatch;
            fallbackLevel = 2;
        }
        else
        {
            uint32_t anyBankPatch = findAnyBank(program);
            if (anyBankPatch != 0xFFFFFFFF)
            {
                // L3: any bank, same program
                finalPatch = anyBankPatch;
                fallbackLevel = 3;
            }
            else
            {
                uint32_t anyDefaultPatch = findAnyBank(0);
                if (anyDefaultPatch != 0xFFFFFFFF)
                {
                    // L4: any bank, program 0
                    finalPatch = anyDefaultPatch;
                    fallbackLevel = 4;
                }
                else if (!m_availablePatches.empty())
                {
                    // L5: absolute fallback, first available instrument
                    finalPatch = *m_availablePatches.begin();
                    fallbackLevel = 5;
                }
            }
        }

        if (fallbackLevel > 0)
        {
            uint32_t fMSB = (finalPatch >> 16) & 0x7F;
            uint32_t fLSB = (finalPatch >> 8) & 0xFF;
            uint32_t fProg = finalPatch & 0xFF;

            EnqueueTranslateLine(
                L"[Translate] Program Change ch%u: Bank %u/%u Prog %u → Bank %u/%u Prog %u (fallback L%d)", channel + 1,
                bankMSB, bankLSB, program, fMSB, fLSB, fProg, fallbackLevel);

            // Re-route to fallback:
            // 1. Send Bank Select
            PackStructuredWithRetry(timestamp, channelGroup, (0xB0 | channel) | (0x00 << 8) | (fMSB << 16));
            PackStructuredWithRetry(timestamp, channelGroup, (0xB0 | channel) | (0x20 << 8) | (fLSB << 16));

            // 2. Send Program Change
            PackStructuredWithRetry(timestamp, channelGroup, (0xC0 | channel) | (fProg << 8));

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

    AccumulateTimestampStats(timestamp, false);

    HRESULT hr = PackStructuredWithRetry(timestamp, channelGroup, midiMessage);
    if (FAILED(hr))
        return false;

    hr = m_pPort->PlayBuffer(m_pMusicBuffer);
    if (FAILED(hr))
        return false;

    m_pMusicBuffer->Flush(); // reset buffer for next message
    return true;
}

bool DmSynth::SendSysEx(const uint8_t* data, uint32_t length, uint32_t channelGroup, REFERENCE_TIME timestamp)
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
            EnqueueTranslateLine(L"[Translate] GM System On → GS Reset");
        else if (isGmOff)
            EnqueueTranslateLine(L"[Translate] GM System Off → GS Reset");
        else if (isXgOn)
            EnqueueTranslateLine(L"[Translate] XG System On → GS Reset");
        static const uint8_t gsResetMsg[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x00, 0x7F, 0x00, 0x41, 0xF7};
        PackUnstructuredWithRetry(timestamp, channelGroup, sizeof(gsResetMsg), gsResetMsg);
        m_pPort->PlayBuffer(m_pMusicBuffer);
        m_pMusicBuffer->Flush();

        ResetInternalState();
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
            EnqueueTranslateLine(L"[Translate] XG Reverb %02X/%02X → GS Reverb macro %02X", xgMsb, xgLsb, gsMacro);
            SendGsReverbMacro(gsMacro);
            return true;
        }
        if (addrLow == 0x20)
        { // Chorus Type MSB (+ optional LSB)
            uint8_t xgMsb = data[7];
            uint8_t xgLsb = (length >= 10) ? data[8] : 0;
            uint8_t gsMacro = MapXgChorusToGs(xgMsb, xgLsb);
            EnqueueTranslateLine(L"[Translate] XG Chorus %02X/%02X → GS Chorus macro %02X", xgMsb, xgLsb, gsMacro);
            SendGsChorusMacro(gsMacro);
            return true;
        }
    }

    AccumulateTimestampStats(timestamp, true);

    HRESULT hr = PackUnstructuredWithRetry(timestamp, channelGroup, length, data);
    if (FAILED(hr))
        return false;

    hr = m_pPort->PlayBuffer(m_pMusicBuffer);
    if (FAILED(hr))
        return false;

    m_pMusicBuffer->Flush();
    return true;
}

void DmSynth::AllSoundOffLocked()
{
    if (!m_pPort || !m_pMusicBuffer)
        return;

    uint32_t groups = m_activeConfig.channelGroups > 0 ? m_activeConfig.channelGroups : 1;
    for (uint32_t group = 1; group <= groups; group++)
    {
        for (uint8_t ch = 0; ch < 16; ch++)
            PackStructuredWithRetry(0, group, (0xB0 | ch) | (120 << 8) | (0 << 16));
        m_pPort->PlayBuffer(m_pMusicBuffer);
        m_pMusicBuffer->Flush();
    }
}

void DmSynth::AllSoundOff()
{
    if (!m_pPort || !m_pMusicBuffer)
        return;
    std::lock_guard<std::mutex> lock(m_mutex);
    AllSoundOffLocked();
}

void DmSynth::ResetAllState()
{
    if (!m_pPort || !m_pMusicBuffer)
        return;
    std::lock_guard<std::mutex> lock(m_mutex);

    // 1. All Sound Off on every channel to silence voices immediately
    AllSoundOffLocked();

    // 2. GS Reset (F0 41 10 42 12 40 00 7F 00 41 F7)
    //    Resets all GS parameters: tuning, pitch bend, controllers, programs, effects, etc.
    static const uint8_t gsResetMsg[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x00, 0x7F, 0x00, 0x41, 0xF7};
    PackUnstructuredWithRetry(0, 1, sizeof(gsResetMsg), gsResetMsg);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();

    // 3. Reset internal tracking state + re-apply custom defaults
    ResetInternalState();
}

// Send GS Reverb Macro SysEx: F0 41 10 42 12 40 01 30 <type> <checksum> F7
// Caller must hold m_mutex.
void DmSynth::SendGsReverbMacro(uint8_t type)
{
    uint8_t sum = 0x40 + 0x01 + 0x30 + type;
    uint8_t checksum = (128 - (sum & 0x7F)) & 0x7F;
    uint8_t sysex[] = {0xF0, 0x41, 0x10, 0x42, 0x12, 0x40, 0x01, 0x30, type, checksum, 0xF7};
    PackUnstructuredWithRetry(0, 1, sizeof(sysex), sysex);
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
    PackUnstructuredWithRetry(0, 1, sizeof(sysex), sysex);
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
    PackUnstructuredWithRetry(0, 1, sizeof(sysex), sysex);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();
}

// Reset internal tracking state and re-apply custom defaults. Caller must hold m_mutex.
void DmSynth::ResetInternalState()
{
    memset(m_bankMSB, 0, sizeof(m_bankMSB));
    memset(m_bankLSB, 0, sizeof(m_bankLSB));
    memset(m_drumChannel, 0, sizeof(m_drumChannel));
    m_drumChannel[9] = true;  // Group 1, CH10
    m_drumChannel[25] = true; // Group 2, CH10
    SendDefaultEffects();
}

// Set default GS effect types. Caller must hold m_mutex.
void DmSynth::SendDefaultEffects()
{
    // GM Master Volume = 100 (Universal Real Time SysEx: F0 7F 7F 04 01 <LSB> <MSB> F7)
    static const uint8_t masterVolMsg[] = {0xF0, 0x7F, 0x7F, 0x04, 0x01, 0x00, 0x64, 0xF7};
    PackUnstructuredWithRetry(0, 1, sizeof(masterVolMsg), masterVolMsg);
    m_pPort->PlayBuffer(m_pMusicBuffer);
    m_pMusicBuffer->Flush();

    SendGsReverbMacro(0x04); // Hall2 (GS default)
    SendGsChorusMacro(0x02); // Chorus3 (GS default)
    SendGsDelayMacro(0x00);  // Delay1 (GS default)
}

bool DmSynth::TryDrainStatsLine(wchar_t* out, size_t cap)
{
    if (!out || cap == 0)
        return false;
    std::lock_guard<std::mutex> slock(m_outLinesMutex);
    if (!m_statsLineReady)
        return false;
    wcsncpy_s(out, cap, m_statsLineBuf, _TRUNCATE);
    m_statsLineReady = false;
    return true;
}

bool DmSynth::TryDrainTranslateLine(wchar_t* out, size_t cap)
{
    if (!out || cap == 0)
        return false;
    std::lock_guard<std::mutex> slock(m_outLinesMutex);
    if (m_pendingTranslateLines.empty())
        return false;
    wcsncpy_s(out, cap, m_pendingTranslateLines.front().c_str(), _TRUNCATE);
    m_pendingTranslateLines.pop_front();
    return true;
}

void DmSynth::EnqueueTranslateLine(const wchar_t* fmt, ...)
{
    wchar_t buf[512];
    va_list ap;
    va_start(ap, fmt);
    vswprintf_s(buf, _countof(buf), fmt, ap);
    va_end(ap);

    std::lock_guard<std::mutex> slock(m_outLinesMutex);
    if (m_pendingTranslateLines.size() >= kMaxPendingTranslateLines)
        m_pendingTranslateLines.pop_front(); // drop oldest on overflow
    m_pendingTranslateLines.emplace_back(buf);
}

void DmSynth::Shutdown()
{
    if (!m_initialized)
        return;

    // Silence any voices still sounding before tearing the port down.
    if (m_pPort && m_pMusicBuffer)
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        AllSoundOffLocked();
    }

    // Give the synth time to render the All Sound Off so trailing audio
    // doesn't bleed into the next session.
    Sleep(500);

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

    if (m_pLatencyClock)
    {
        m_pLatencyClock->Release();
        m_pLatencyClock = nullptr;
    }
    m_anchorInitialized = false;

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

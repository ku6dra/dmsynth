#pragma once

#include <deque>
#include <string>
#include <vector>
#include <unordered_set>
#include <mutex>
#include <windows.h>
#include <dsound.h>
#include "directmusic_defs.h"

// Set to 1 to log every timestamp handed to IDirectMusicBuffer. The guarded
// code is preprocessed out entirely when 0, so Release builds carry no cost.
#ifndef DMSYNTH_DEBUG_TIMESTAMPS
#define DMSYNTH_DEBUG_TIMESTAMPS 0
#endif

struct SynthPortInfo
{
    uint32_t index;
    GUID guid;
    std::wstring description;
    uint32_t dwClass;
    uint32_t maxVoices;
    uint32_t maxChannelGroups;
    uint32_t flags;
    bool isSoftwareSynth;
};

struct SynthConfig
{
    uint32_t sampleRate = 44100; // DEFAULT: 44.1kHz Stereo for Native DirectMusic
    uint32_t voices = 128;
    uint32_t channelGroups = 2; // 2 channel groups = 32 MIDI channels
    uint32_t audioChannels = 2; // Stereo
    std::wstring dlsPath;       // DLS file path (empty = system gm.dls)
};

class DmSynth
{
  public:
    DmSynth();
    ~DmSynth();

    // Prevent copying
    DmSynth(const DmSynth&) = delete;
    DmSynth& operator=(const DmSynth&) = delete;

    bool Initialize();
    std::vector<SynthPortInfo> EnumeratePorts();
    bool CreateSynthPort(const GUID& portGuid, const SynthConfig& config);
    bool SendMidiMessage(uint8_t status, uint8_t data1, uint8_t data2, uint32_t channelGroup = 1,
                         REFERENCE_TIME timestamp = 0);
    bool SendSysEx(const uint8_t* data, uint32_t length, uint32_t channelGroup = 1, REFERENCE_TIME timestamp = 0);
    void ResetAllState();

    // Send CC 120 (All Sound Off) to every channel in every active group.
    // Safe to call from shutdown paths to prevent hung voices.
    void AllSoundOff();

    // Convert an absolute timeGetTime() millisecond value (i.e. a MIDI
    // arrival time already rebased onto timeGetTime's domain via
    // MidiInput::GetStartTimeMs()) to a DirectMusic REFERENCE_TIME on the
    // port's latency clock, with a small scheduling lead so messages are
    // rendered at a deterministic future instant rather than "as soon as
    // they arrive" — this absorbs callback jitter.
    REFERENCE_TIME MidiMsToRefTime(DWORD absWinmmMs);

    const SynthConfig& GetActiveConfig() const { return m_activeConfig; }
    void Shutdown();

    // Non-blocking drain of the once-per-second stats line. Called from a
    // non-callback thread so the actual console write (which can stall
    // under ConPTY/slow terminals) never blocks the MIDI callback thread.
    // Returns true if a line was pending; writes a NUL-terminated string
    // (without trailing newline) into `out`.
    bool TryDrainStatsLine(wchar_t* out, size_t cap);

    // Non-blocking drain of one queued [Translate ...] line, produced on
    // the callback thread when GS/XG/GM translation fires. Returns true
    // if a line was dequeued. Same rationale as TryDrainStatsLine: console
    // writes never happen on the callback thread.
    bool TryDrainTranslateLine(wchar_t* out, size_t cap);

  private:
    bool DownloadGMInstruments();
    void SendGsReverbMacro(uint8_t type);
    void SendGsChorusMacro(uint8_t type);
    void SendGsDelayMacro(uint8_t type);
    void SendDefaultEffects();
    void ResetInternalState();
    void AllSoundOffLocked(); // caller holds m_mutex

    // PackStructured/PackUnstructured wrappers that flush+retry once on
    // failure (e.g. DMUS_E_BUFFER_FULL) so a transient full buffer doesn't
    // drop MIDI data. Caller must hold m_mutex.
    HRESULT PackStructuredWithRetry(REFERENCE_TIME rt, uint32_t group, DWORD msg);
    HRESULT PackUnstructuredWithRetry(REFERENCE_TIME rt, uint32_t group, uint32_t length, const uint8_t* data);

    // Accumulate scheduling-lead stats for a single outbound message and,
    // roughly once per second, run the auto-tune controller that nudges
    // m_schedulingLead up when messages arrive late and down when there is
    // excess headroom. Always compiled so the controller runs in Release
    // builds; only the per-window fwprintf is gated by DMSYNTH_DEBUG_TIMESTAMPS.
    // Caller holds m_mutex.
    void AccumulateTimestampStats(REFERENCE_TIME scheduledRt, bool isSysEx);

    struct TimestampStats
    {
        REFERENCE_TIME windowStart = 0; // latency-clock time window began
        uint32_t midiCount = 0;
        uint32_t sysexCount = 0;
        uint32_t lateCount = 0;  // lead < 0 (DirectMusic sees past)
        uint32_t retryCount = 0; // buffer-full retries
        REFERENCE_TIME leadMin = 0;
        REFERENCE_TIME leadMax = 0;
        REFERENCE_TIME leadSum = 0;
        // Callback delay = handler time - MIDI arrival time. Derived directly
        // as (schedulingLead - lead), so it shows how jittery the MIDI
        // callback itself is — which is the real knob for tuning lead.
        REFERENCE_TIME cbDelayMin = 0;
        REFERENCE_TIME cbDelayMax = 0;
        REFERENCE_TIME cbDelaySum = 0;
        bool haveSample = false;
    };
    TimestampStats m_tsStats;

    IDirectMusic8* m_pDirectMusic = nullptr;
    IDirectSound8* m_pDirectSound = nullptr;
    IDirectMusicPort* m_pPort = nullptr;
    IDirectMusicBuffer* m_pMusicBuffer = nullptr;
    IReferenceClock* m_pLatencyClock = nullptr;

    // Anchor pair: (timeGetTime ms, latency-clock REFERENCE_TIME) sampled
    // together on the MIDI callback thread at the very first message and
    // held for the port's lifetime. Using timeGetTime (not the arrival-time
    // MIDI stamp) means callback-delay jitter is not baked into the anchor.
    // The anchor's one-time ε becomes a constant per-session offset rather
    // than a periodic audible step.
    REFERENCE_TIME m_refAnchor = 0;
    DWORD m_winmmAnchorMs = 0;
    bool m_anchorInitialized = false;
    // Added to each outgoing timestamp. Signed so the controller can go
    // below 0 when anchor ε is large (schedule in the past → DirectMusic
    // renders "as soon as possible"). Auto-tuned at runtime; start value
    // is just a plausible seed.
    REFERENCE_TIME m_schedulingLead = 100000; // 10 ms seed

    // Auto-tune controller state. Inspects each 1-second stats window and
    // nudges m_schedulingLead by 0.25 ms (down) up to ~2.0 ms (accelerating
    // up-step, see kAutoTuneUpShiftMax) to keep actualLeadMin (= lead -
    // cbDelayMax) inside a fixed band. Frozen once headroom has sat in
    // band for long enough; unfreezes on any out-of-band signal.
    // There is no cap on lead itself — capping it would let anchor-ε
    // starve the real-world slack.
    bool m_autoTuneFrozen = false;
    uint32_t m_autoTuneStableSeconds = 0;
    uint32_t m_autoTuneConsecUp = 0; // consecutive +adjust actions for asymmetric acceleration
    int32_t m_autoTunePending = 0;   // signed vote accumulator: +n wants up, -n wants down
    static constexpr REFERENCE_TIME kAutoTuneMinLeadRt = -100 * 10000; // allow "past" scheduling
    static constexpr REFERENCE_TIME kAutoTuneStepRt = 2500;            // 0.25 ms base step
    static constexpr uint32_t kAutoTuneUpShiftMax = 3;                 // up-step doubles up to 3x -> 2.0 ms cap
    static constexpr int32_t kAutoTuneHystVotes = 2;                   // out-of-band windows needed to trigger
    static constexpr uint32_t kAutoTuneMinSamples = 30;                // min msgs per eval window; carry over if short
    static constexpr double kAutoTuneTargetLowMs = 1.0;  // actualLeadMin floor
    static constexpr double kAutoTuneTargetHighMs = 7.0; // actualLeadMin ceiling
    static constexpr uint32_t kAutoTuneFreezeSeconds = 30;

    // DLS Management
    IDirectMusicLoader8* m_pLoader = nullptr;
    IDirectMusicCollection8* m_pDefaultDls = nullptr;
    std::vector<IDirectMusicDownloadedInstrument*> m_downloadedInstruments;
    std::unordered_set<DWORD> m_availablePatches;

    // MIDI State for fallback logic
    uint8_t m_bankMSB[32] = {0}; // Track MSB for each channel (up to 2 groups)
    uint8_t m_bankLSB[32] = {0}; // Track LSB for each channel
    bool m_drumChannel[32] = {}; // Track which channels are drum channels

    SynthConfig m_activeConfig;
    std::mutex m_mutex;
    bool m_initialized = false;
    bool m_comInitialized = false;

    // Guards both the pending stats line (m_statsLineBuf/Ready) and the
    // translate-line queue. Producers are the callback thread; consumer
    // is the logger thread via TryDrain*. Keeps console writes off the
    // callback thread.
    std::mutex m_outLinesMutex;
    wchar_t m_statsLineBuf[512] = {};
    bool m_statsLineReady = false;

    // Capped so a runaway translation burst cannot grow memory unbounded.
    std::deque<std::wstring> m_pendingTranslateLines;
    static constexpr size_t kMaxPendingTranslateLines = 128;
    void EnqueueTranslateLine(const wchar_t* fmt, ...);
};

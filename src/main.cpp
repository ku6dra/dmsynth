#include "cli_args.h"
#include "dm_synth.h"
#include "midi_input.h"
#include <atomic>
#include <cstdint>
#include <conio.h>
#include <cstdio>
#include <cstdlib>
#include <fcntl.h>
#include <io.h>
#include <thread>

static std::atomic<bool> g_running{true};
static std::atomic<bool> g_verboseLog{false};

// Signaled by the main thread after Shutdown() completes. The CTRL_CLOSE/
// LOGOFF/SHUTDOWN handler waits on this before returning — otherwise Windows
// calls ExitProcess as soon as the handler returns and kills the main thread
// mid-cleanup, leaving DirectMusic/DirectSound holding ~100 ms of audio that
// then bleeds into the next session's startup.
static HANDLE g_shutdownComplete = nullptr;

// --- Async MIDI Log (SPSC ring buffer + event) ---
// Callback thread pushes raw events; main loop wakes on signal and prints.
struct MidiLogEvent
{
    uint8_t kind; // 0=short, 1=sysex
    uint8_t status;
    uint8_t data1;
    uint8_t data2;
    uint8_t group; // channel group: 1 or 2
    DWORD sysExLength;
};

static constexpr size_t LOG_RING_SIZE = 2048;
static MidiLogEvent g_logRing[LOG_RING_SIZE];
static std::atomic<size_t> g_logHead{0};
static std::atomic<size_t> g_logTail{0};
static HANDLE g_logEvent = nullptr; // auto-reset event

static void LogPush(const MidiLogEvent& evt)
{
    if (!g_verboseLog.load(std::memory_order_relaxed))
        return;
    size_t head = g_logHead.load(std::memory_order_relaxed);
    size_t next = (head + 1) % LOG_RING_SIZE;
    if (next == g_logTail.load(std::memory_order_acquire))
        return; // full — drop the log entry (audio is still processed)
    g_logRing[head] = evt;
    g_logHead.store(next, std::memory_order_release);
    SetEvent(g_logEvent); // wake the main loop immediately
}

static bool LogPop(MidiLogEvent& evt)
{
    size_t tail = g_logTail.load(std::memory_order_relaxed);
    if (tail == g_logHead.load(std::memory_order_acquire))
        return false;
    evt = g_logRing[tail];
    g_logTail.store((tail + 1) % LOG_RING_SIZE, std::memory_order_release);
    return true;
}

// Read one line from stdin and parse an integer.
// Returns true and sets `out` if a valid integer was found.
// Returns false (empty line / non-numeric) so callers can fall back to a default.
static bool ReadLineInt(int& out)
{
    wchar_t buf[64] = {};
    if (!fgetws(buf, _countof(buf), stdin))
        return false;
    wchar_t* end = nullptr;
    long val = wcstol(buf, &end, 10);
    if (end == buf) // nothing was parsed (e.g. bare Enter)
        return false;
    out = static_cast<int>(val);
    return true;
}

static BOOL WINAPI ConsoleCtrlHandler(DWORD ctrlType)
{
    switch (ctrlType)
    {
    case CTRL_C_EVENT:
    case CTRL_BREAK_EVENT:
        g_running = false;
        return TRUE;
    case CTRL_CLOSE_EVENT:
    case CTRL_LOGOFF_EVENT:
    case CTRL_SHUTDOWN_EVENT:
        g_running = false;
        // After this handler returns, the OS calls ExitProcess and terminates
        // every other thread — block here until the main thread finishes
        // synth.Shutdown(), so DirectMusic/DirectSound is properly torn down.
        if (g_shutdownComplete)
            WaitForSingleObject(g_shutdownComplete, 4000);
        return TRUE;
    default:
        return FALSE;
    }
}

static const wchar_t* NoteName(uint8_t note)
{
    static const wchar_t* names[] = {L"C", L"C#", L"D", L"D#", L"E", L"F", L"F#", L"G", L"G#", L"A", L"A#", L"B"};
    return names[note % 12];
}

static int NoteOctave(uint8_t note)
{
    return static_cast<int>(note / 12) - 1;
}

int wmain(int argc, wchar_t* argv[])
{
    // --- Parse command-line arguments ---
    auto args = ParseCommandLine(argc, argv);

    // Set console to UTF-16 output
    _setmode(_fileno(stdout), _O_U16TEXT);
    _setmode(_fileno(stderr), _O_U16TEXT);

    // Enable ANSI escape sequences for colored output
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    DWORD consoleMode = 0;
    if (GetConsoleMode(hConsole, &consoleMode))
    {
        SetConsoleMode(hConsole, consoleMode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
    }

    if (args.parseError)
    {
        fwprintf(stderr, L"Error: %s\n\n", args.errorMessage.c_str());
        PrintUsage(argv[0]);
        return 1;
    }
    if (args.showHelp)
    {
        PrintUsage(argv[0]);
        return 0;
    }

    g_verboseLog.store(args.verbose, std::memory_order_relaxed);

    g_shutdownComplete = CreateEventW(nullptr, TRUE, FALSE, nullptr); // manual-reset
    SetConsoleCtrlHandler(ConsoleCtrlHandler, TRUE);

    wprintf(L"==============================================\n");
    wprintf(L"  DMSynth - DirectMusic MIDI Synthesizer\n");
    wprintf(L"==============================================\n\n");

    // --- Initialize DirectMusic Synth ---
    DmSynth synth;
    if (!synth.Initialize())
    {
        fwprintf(stderr, L"Failed to initialize DirectMusic.\n");
        return 1;
    }

    // --- Enumerate and select synth port ---
    auto ports = synth.EnumeratePorts();
    if (ports.empty())
    {
        fwprintf(stderr, L"No synthesizer ports available.\n");
        return 1;
    }

    // --- --list: show devices and exit ---
    if (args.listDevices)
    {
        wprintf(L"--- Synthesizer Ports ---\n");
        for (size_t i = 0; i < ports.size(); i++)
        {
            wprintf(L"  [%zu] %s%s\n", i, ports[i].description.c_str(),
                    ports[i].isSoftwareSynth ? L" (Software)" : L"");
        }
        wprintf(L"\n");
        auto listMidiDevices = MidiInput::EnumerateDevices();
        if (listMidiDevices.empty())
        {
            wprintf(L"--- MIDI Input Devices ---\n");
            wprintf(L"  (none)\n");
        }
        else
        {
            wprintf(L"--- MIDI Input Devices ---\n");
            for (size_t i = 0; i < listMidiDevices.size(); i++)
            {
                wprintf(L"  [%zu] %s\n", i, listMidiDevices[i].name.c_str());
            }
        }
        synth.Shutdown();
        return 0;
    }

    // --- Select synth port ---
    int portChoice = -1;
    if (args.port >= 0)
    {
        if (args.port >= static_cast<int>(ports.size()))
        {
            fwprintf(stderr, L"Error: port index %d out of range (0-%zu)\n", args.port, ports.size() - 1);
            synth.Shutdown();
            return 1;
        }
        portChoice = args.port;
    }
    else
    {
        // Auto-select the software synthesizer port (Microsoft Synthesizer)
        for (size_t i = 0; i < ports.size(); i++)
        {
            if (ports[i].isSoftwareSynth)
            {
                portChoice = static_cast<int>(i);
                break;
            }
        }
        if (portChoice < 0)
        {
            portChoice = 0; // fallback to first port
        }
    }
    wprintf(L"Synthesizer: %s\n\n", ports[portChoice].description.c_str());

    // --- Configure and create synth port ---
    SynthConfig config;
    config.sampleRate = (args.sampleRate > 0) ? static_cast<uint32_t>(args.sampleRate) : 44100;
    config.voices = (args.voices > 0) ? static_cast<uint32_t>(args.voices) : 128;
    config.channelGroups = 2;
    config.audioChannels = 2;
    config.dlsPath = args.dlsPath;
    config.forceImmediateMode = args.immediate;

    if (!synth.CreateSynthPort(ports[portChoice].guid, config))
    {
        fwprintf(stderr, L"Failed to create synthesizer port.\n");
        return 1;
    }

    auto activeConfig = synth.GetActiveConfig();
    wprintf(L"--- Synthesizer Settings ---\n");
    wprintf(L"  Sample rate:     %lu Hz\n", activeConfig.sampleRate);
    wprintf(L"  Voices:          %lu\n", activeConfig.voices);
    wprintf(L"  Channel groups:  %lu\n", activeConfig.channelGroups);
    wprintf(L"  Audio channels:  %lu\n", activeConfig.audioChannels);
    wprintf(L"\n");

    // --- Enumerate and select MIDI input device ---
    auto midiDevices = MidiInput::EnumerateDevices();
    if (midiDevices.empty())
    {
        fwprintf(stderr, L"No MIDI input devices found.\n");
        fwprintf(stderr, L"Please connect a MIDI keyboard or controller.\n");
        return 1;
    }

    wprintf(L"--- MIDI Input Devices ---\n");
    for (size_t i = 0; i < midiDevices.size(); i++)
    {
        wprintf(L"  [%zu] %s\n", i, midiDevices[i].name.c_str());
    }

    // --- MIDI IN 1: CH1-16 (Channel Group 1) ---
    int midiChoice1 = 0;
    if (args.midi1.specified)
    {
        std::wstring resolveErr;
        midiChoice1 = ResolveMidiDevice(args.midi1, midiDevices, L"MIDI IN 1", resolveErr);
        if (midiChoice1 < 0)
        {
            fwprintf(stderr, L"Error: %s\n", resolveErr.c_str());
            return 1;
        }
    }
    else if (midiDevices.size() > 1)
    {
        wprintf(L"\nSelect MIDI IN 1 (CH1-16) [0-%zu]: ", midiDevices.size() - 1);
        fflush(stdout);
        if (!ReadLineInt(midiChoice1) || midiChoice1 < 0 || midiChoice1 >= static_cast<int>(midiDevices.size()))
        {
            midiChoice1 = 0;
        }
    }
    wprintf(L"\nMIDI IN 1: %s\n", midiDevices[midiChoice1].name.c_str());

    // --- MIDI IN 2: CH17-32 (Channel Group 2) ---
    int midiChoice2 = -1;
    if (args.midi2.specified)
    {
        if (args.midi2.isNone)
        {
            midiChoice2 = -1;
        }
        else
        {
            std::wstring resolveErr;
            midiChoice2 = ResolveMidiDevice(args.midi2, midiDevices, L"MIDI IN 2", resolveErr);
            if (midiChoice2 < 0)
            {
                fwprintf(stderr, L"Error: %s\n", resolveErr.c_str());
                return 1;
            }
        }
    }
    else if (static_cast<int>(midiDevices.size()) >= 2)
    {
        wprintf(L"Select MIDI IN 2 (CH17-32) [0-%zu, -1=none]: ", midiDevices.size() - 1);
        fflush(stdout);
        if (!ReadLineInt(midiChoice2) || midiChoice2 < -1 || midiChoice2 >= static_cast<int>(midiDevices.size()))
        {
            midiChoice2 = -1;
        }
    }
    if (midiChoice2 >= 0 && midiChoice2 == midiChoice1)
    {
        wprintf(L"Cannot use the same device as MIDI IN 1. MIDI IN 2: none\n\n");
        midiChoice2 = -1;
    }
    else if (midiChoice2 >= 0)
        wprintf(L"MIDI IN 2: %s\n\n", midiDevices[midiChoice2].name.c_str());
    else
        wprintf(L"MIDI IN 2: none\n\n");

    // --- Set up MIDI input with routing to synth ---
    timeBeginPeriod(1);
    MidiInput midiIn1, midiIn2;
    g_logEvent = CreateEventW(nullptr, FALSE, FALSE, nullptr); // auto-reset

    // Port 1: CH1-16 → Channel Group 1
    // NOTE: This callback runs on the multimedia timer thread.
    //       Only push to the ring buffer here — no I/O allowed.
    midiIn1.SetCallback([&synth, &midiIn1](uint8_t status, uint8_t data1, uint8_t data2, DWORD midiMs) {
        DWORD absMs = midiIn1.GetStartTimeMs() + midiMs;
        REFERENCE_TIME rt = synth.MidiMsToRefTime(absMs);
        synth.SendMidiMessage(status, data1, data2, 1, rt);
        LogPush({0, status, data1, data2, 1, 0});
    });
    midiIn1.SetSysExCallback([&synth, &midiIn1](const uint8_t* data, DWORD length, DWORD midiMs) {
        DWORD absMs = midiIn1.GetStartTimeMs() + midiMs;
        REFERENCE_TIME rt = synth.MidiMsToRefTime(absMs);
        synth.SendSysEx(data, length, 1, rt);
        LogPush({1, 0, 0, 0, 1, length});
    });

    if (!midiIn1.Open(midiDevices[midiChoice1].id))
    {
        fwprintf(stderr, L"Failed to open MIDI IN 1.\n");
        CloseHandle(g_logEvent);
        return 1;
    }

    // Port 2: CH17-32 → Channel Group 2 (optional)
    if (midiChoice2 >= 0)
    {
        midiIn2.SetCallback([&synth, &midiIn2](uint8_t status, uint8_t data1, uint8_t data2, DWORD midiMs) {
            DWORD absMs = midiIn2.GetStartTimeMs() + midiMs;
            REFERENCE_TIME rt = synth.MidiMsToRefTime(absMs);
            synth.SendMidiMessage(status, data1, data2, 2, rt);
            LogPush({0, status, data1, data2, 2, 0});
        });
        midiIn2.SetSysExCallback([&synth, &midiIn2](const uint8_t* data, DWORD length, DWORD midiMs) {
            DWORD absMs = midiIn2.GetStartTimeMs() + midiMs;
            REFERENCE_TIME rt = synth.MidiMsToRefTime(absMs);
            synth.SendSysEx(data, length, 2, rt);
            LogPush({1, 0, 0, 0, 2, length});
        });
        if (!midiIn2.Open(midiDevices[midiChoice2].id))
            fwprintf(stderr, L"Failed to open MIDI IN 2. (continuing)\n");
    }

    wprintf(L"==============================================\n");
    wprintf(L"  Receiving MIDI input... Ctrl+C to exit\n");
    wprintf(L"  R = Reset / V = Toggle MIDI log\n");
    wprintf(L"==============================================\n\n");

    // Dedicated logger thread: wprintf/fflush happen here so the main thread
    // (which handles input and window messages) isn't blocked by slow
    // terminals (ConPTY char-by-char rendering can stall for tens of ms).
    std::thread logThread([&synth] {
        wchar_t statsLine[512];
        while (g_running)
        {
            WaitForSingleObject(g_logEvent, 100);

            // Drain any pending per-second stats line produced on the MIDI
            // callback thread. Printed here so the actual console write
            // (ConPTY/terminal stalls) never blocks the callback thread.
            if (synth.TryDrainStatsLine(statsLine, _countof(statsLine)))
            {
                fwprintf(stderr, L"%s\n", statsLine);
                fflush(stderr);
            }

            // Same rationale for queued [Translate ...] lines from GS/XG/GM
            // translation paths.
            bool wroteTranslate = false;
            while (synth.TryDrainTranslateLine(statsLine, _countof(statsLine)))
            {
                wprintf(L"%s\n", statsLine);
                wroteTranslate = true;
            }
            if (wroteTranslate)
                fflush(stdout);

            MidiLogEvent evt;
            bool wrote = false;
            while (LogPop(evt))
            {
                if (!g_verboseLog.load(std::memory_order_relaxed))
                    continue;
                wrote = true;
                if (evt.kind == 1)
                {
                    wprintf(L"  \x1b[95m[G%u] SysEx: %lu bytes\x1b[0m\n", evt.group, evt.sysExLength);
                    continue;
                }
                uint8_t msgType = evt.status & 0xF0;
                uint8_t channel = evt.status & 0x0F;
                uint8_t displayCh = channel + 1 + (evt.group - 1) * 16;
                switch (msgType)
                {
                case 0x90:
                    if (evt.data2 > 0)
                        wprintf(L"  \x1b[92m\u266a Note ON  ch=%2u  %s%d  vel=%3u\x1b[0m\n", displayCh,
                                NoteName(evt.data1), NoteOctave(evt.data1), evt.data2);
                    else
                        wprintf(L"  \x1b[90m\u266a Note OFF ch=%2u  %s%d\x1b[0m\n", displayCh, NoteName(evt.data1),
                                NoteOctave(evt.data1));
                    break;
                case 0x80:
                    wprintf(L"  \x1b[90m\u266a Note OFF ch=%2u  %s%d\x1b[0m\n", displayCh, NoteName(evt.data1),
                            NoteOctave(evt.data1));
                    break;
                case 0xB0:
                    wprintf(L"  \x1b[33mCC ch=%2u  cc=%3u  val=%3u\x1b[0m\n", displayCh, evt.data1, evt.data2);
                    break;
                case 0xC0:
                    wprintf(L"  \x1b[36mPC ch=%2u  prog=%3u\x1b[0m\n", displayCh, evt.data1);
                    break;
                case 0xE0: {
                    int bend = (static_cast<int>(evt.data2) << 7) | evt.data1;
                    wprintf(L"  \x1b[35mPB ch=%2u  bend=%5d\x1b[0m\n", displayCh, bend - 8192);
                }
                break;
                case 0xD0:
                    wprintf(L"  \x1b[34mCP ch=%2u  pressure=%3u\x1b[0m\n", displayCh, evt.data1);
                    break;
                }
            }
            if (wrote)
                fflush(stdout);
        }
    });

    // Main loop - handles keyboard input and window messages only.
    while (g_running)
    {
        MsgWaitForMultipleObjects(0, nullptr, FALSE, 50, QS_ALLEVENTS);

        while (_kbhit())
        {
            int key = _getwch();
            if (key == L'r' || key == L'R')
            {
                synth.ResetAllState();
                wprintf(L"\n  ** Reset complete **\n\n");
                fflush(stdout);
            }
            else if (key == L'v' || key == L'V')
            {
                bool newVal = !g_verboseLog.load(std::memory_order_relaxed);
                g_verboseLog.store(newVal, std::memory_order_relaxed);
                wprintf(L"\n  ** MIDI log: %s **\n\n", newVal ? L"ON" : L"OFF");
                fflush(stdout);
            }
        }

        MSG msg;
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    SetEvent(g_logEvent); // wake logger so it sees g_running=false
    logThread.join();

    wprintf(L"\nShutting down...\n");

    midiIn1.Close();
    midiIn2.Close();
    synth.Shutdown();
    CloseHandle(g_logEvent);
    timeEndPeriod(1);

    // Tell the CTRL_CLOSE handler (if any) it can let the OS exit the process.
    if (g_shutdownComplete)
    {
        SetEvent(g_shutdownComplete);
        // Don't CloseHandle: the close-handler thread may still be inside
        // WaitForSingleObject. The OS reclaims it on process exit anyway.
    }

    return 0;
}

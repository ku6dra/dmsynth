#include "cli_args.h"
#include "midi_input.h"
#include <algorithm>
#include <cstdio>
#include <cstdlib>
#include <cwchar>

static bool TryParseInt(const wchar_t* str, int& out)
{
    wchar_t* end = nullptr;
    long val = wcstol(str, &end, 10);
    if (end == str || *end != L'\0')
        return false;
    out = static_cast<int>(val);
    return true;
}

static bool ConsumeIntArg(int argc, wchar_t* argv[], int& i, const wchar_t* name, int& out, std::wstring& err)
{
    if (i + 1 >= argc)
    {
        err = std::wstring(L"Option ") + name + L" requires a value";
        return false;
    }
    ++i;
    if (!TryParseInt(argv[i], out))
    {
        err = std::wstring(L"Option ") + name + L": invalid value: " + argv[i];
        return false;
    }
    return true;
}

static bool ConsumeStringArg(int argc, wchar_t* argv[], int& i, const wchar_t* name, std::wstring& out,
                             std::wstring& err)
{
    if (i + 1 >= argc)
    {
        err = std::wstring(L"Option ") + name + L" requires a value";
        return false;
    }
    ++i;
    out = argv[i];
    return true;
}

CliArgs ParseCommandLine(int argc, wchar_t* argv[])
{
    CliArgs args;

    for (int i = 1; i < argc; ++i)
    {
        std::wstring arg = argv[i];

        if (arg == L"--help" || arg == L"-h")
        {
            args.showHelp = true;
            return args;
        }
        else if (arg == L"--list")
        {
            args.listDevices = true;
            return args;
        }
        else if (arg == L"--verbose" || arg == L"-v")
        {
            args.verbose = true;
        }
        else if (arg == L"--immediate")
        {
            args.immediate = true;
        }
        else if (arg == L"--midi1")
        {
            args.midi1.specified = true;
            if (!ConsumeStringArg(argc, argv, i, L"--midi1", args.midi1.rawValue, args.errorMessage))
            {
                args.parseError = true;
                return args;
            }
        }
        else if (arg == L"--midi2")
        {
            args.midi2.specified = true;
            if (!ConsumeStringArg(argc, argv, i, L"--midi2", args.midi2.rawValue, args.errorMessage))
            {
                args.parseError = true;
                return args;
            }
            // Special case: -1 means "none"
            if (args.midi2.rawValue == L"-1")
            {
                args.midi2.isNone = true;
            }
        }
        else if (arg == L"--port" || arg == L"-p")
        {
            if (!ConsumeIntArg(argc, argv, i, L"--port", args.port, args.errorMessage))
            {
                args.parseError = true;
                return args;
            }
        }
        else if (arg == L"--rate" || arg == L"-r")
        {
            if (!ConsumeIntArg(argc, argv, i, L"--rate", args.sampleRate, args.errorMessage))
            {
                args.parseError = true;
                return args;
            }
            if (args.sampleRate <= 0)
            {
                args.parseError = true;
                args.errorMessage = L"--rate must be a positive integer";
                return args;
            }
        }
        else if (arg == L"--voices")
        {
            if (!ConsumeIntArg(argc, argv, i, L"--voices", args.voices, args.errorMessage))
            {
                args.parseError = true;
                return args;
            }
            if (args.voices <= 0)
            {
                args.parseError = true;
                args.errorMessage = L"--voices must be a positive integer";
                return args;
            }
        }
        else if (arg == L"--dls")
        {
            if (!ConsumeStringArg(argc, argv, i, L"--dls", args.dlsPath, args.errorMessage))
            {
                args.parseError = true;
                return args;
            }
        }
        else
        {
            args.parseError = true;
            args.errorMessage = std::wstring(L"Unknown option: ") + arg;
            return args;
        }
    }

    return args;
}

void PrintUsage(const wchar_t* programName)
{
    wprintf(L"Usage: %s [options]\n", programName);
    wprintf(L"\n");
    wprintf(L"Options:\n");
    wprintf(L"  --midi1 <idx|name> MIDI IN 1 device (CH1-16)\n");
    wprintf(L"                     Specify by index number or partial device name\n");
    wprintf(L"                     Name matching is case-insensitive substring search\n");
    wprintf(L"  --midi2 <idx|name> MIDI IN 2 device (CH17-32)\n");
    wprintf(L"                     Use -1 to disable\n");
    wprintf(L"  --port, -p <idx>   Synthesizer port index\n");
    wprintf(L"  --rate, -r <hz>    Sample rate (Hz, e.g. 44100, 48000)\n");
    wprintf(L"  --voices <n>       Maximum number of voices\n");
    wprintf(L"  --dls <path>       DLS file path (default: system gm.dls)\n");
    wprintf(L"  --verbose, -v      Enable MIDI logging from startup\n");
    wprintf(L"  --immediate        Force rt=0 scheduling (no latency-clock probe)\n");
    wprintf(L"                     Use when running alongside software like DirectMusic Producer\n");
    wprintf(L"  --list             List devices/ports and exit\n");
    wprintf(L"  --help, -h         Show this help and exit\n");
    wprintf(L"\n");
    wprintf(L"Unspecified options use default values or interactive prompts.\n");
    wprintf(L"\n");
    wprintf(L"Examples:\n");
    wprintf(L"  %s --list\n", programName);
    wprintf(L"  %s --midi1 0 --midi2 -1 --rate 44100\n", programName);
    wprintf(L"  %s --midi1 midikeysmasher --midi2 1 -p 0 -v\n", programName);
}

// Case-insensitive substring search for wide strings
static bool ContainsCaseInsensitive(const std::wstring& haystack, const std::wstring& needle)
{
    if (needle.empty())
        return true;
    if (haystack.size() < needle.size())
        return false;

    auto toLower = [](wchar_t c) -> wchar_t { return towlower(c); };

    auto it = std::search(haystack.begin(), haystack.end(), needle.begin(), needle.end(),
                          [&](wchar_t a, wchar_t b) { return toLower(a) == toLower(b); });
    return it != haystack.end();
}

int ResolveMidiDevice(const MidiDeviceSpec& spec, const std::vector<MidiDeviceInfo>& devices, const wchar_t* label,
                      std::wstring& errorMsg)
{
    // Try to parse as integer index first
    int idx = 0;
    if (TryParseInt(spec.rawValue.c_str(), idx))
    {
        if (idx < 0 || idx >= static_cast<int>(devices.size()))
        {
            errorMsg = std::wstring(label) + L" index " + std::to_wstring(idx) + L" out of range (0-" +
                       std::to_wstring(devices.size() - 1) + L")";
            return -1;
        }
        return idx;
    }

    // Partial name match (case-insensitive)
    int matchIndex = -1;
    int matchCount = 0;
    for (size_t i = 0; i < devices.size(); i++)
    {
        if (ContainsCaseInsensitive(devices[i].name, spec.rawValue))
        {
            matchIndex = static_cast<int>(i);
            matchCount++;
        }
    }

    if (matchCount == 0)
    {
        errorMsg = std::wstring(label) + L" \"" + spec.rawValue + L"\": no matching device found";
        return -1;
    }
    if (matchCount > 1)
    {
        errorMsg = std::wstring(label) + L" \"" + spec.rawValue + L"\": " + std::to_wstring(matchCount) +
                   L" devices matched; please use a more specific name";
        return -1;
    }

    return matchIndex;
}

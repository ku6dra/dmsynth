#pragma once

#include <string>
#include <vector>

struct MidiDeviceSpec
{
    bool specified = false; // true if CLI argument was given
    bool isNone = false;    // true if explicitly set to "none" (-1 for midi2)
    std::wstring rawValue;  // raw argument string (index or partial name)
};

struct CliArgs
{
    MidiDeviceSpec midi1;
    MidiDeviceSpec midi2;
    int port = -1;        // -1 = auto-select first software synth
    int sampleRate = -1;  // -1 = use default (44100)
    int voices = -1;      // -1 = use default (128)
    std::wstring dlsPath; // empty = system gm.dls
    bool verbose = false;
    bool listDevices = false;
    bool showHelp = false;
    bool parseError = false;
    std::wstring errorMessage;
};

struct MidiDeviceInfo; // forward declaration

// Parse command-line arguments
CliArgs ParseCommandLine(int argc, wchar_t* argv[]);

// Print usage help
void PrintUsage(const wchar_t* programName);

// Resolve a MidiDeviceSpec to a device index.
// Returns the matched device index, or -1 on error (with errorMsg set).
int ResolveMidiDevice(const MidiDeviceSpec& spec, const std::vector<MidiDeviceInfo>& devices, const wchar_t* label,
                      std::wstring& errorMsg);

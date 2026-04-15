#pragma once

#include <windows.h>
#include <mmsystem.h>
#include <avrt.h>
#include <atomic>
#include <functional>
#include <string>
#include <vector>
#include <cstdint>

struct MidiDeviceInfo
{
    UINT id;
    std::wstring name;
    WORD mid; // manufacturer ID
    WORD pid; // product ID
};

// Callback: status byte, data1, data2, timestamp
using MidiCallback = std::function<void(uint8_t status, uint8_t data1, uint8_t data2, DWORD timestamp)>;
// SysEx callback: data pointer, length, timestamp
using MidiSysExCallback = std::function<void(const uint8_t* data, DWORD length, DWORD timestamp)>;

class MidiInput
{
  public:
    MidiInput();
    ~MidiInput();

    MidiInput(const MidiInput&) = delete;
    MidiInput& operator=(const MidiInput&) = delete;

    static std::vector<MidiDeviceInfo> EnumerateDevices();

    bool Open(UINT deviceId);
    void Close();
    bool IsOpen() const { return m_hMidiIn != nullptr; }

    void SetCallback(MidiCallback cb) { m_callback = std::move(cb); }
    void SetSysExCallback(MidiSysExCallback cb) { m_sysExCallback = std::move(cb); }

  private:
    static void CALLBACK MidiInProc(HMIDIIN hMidiIn, UINT wMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1,
                                    DWORD_PTR dwParam2);

    void HandleShortMessage(DWORD_PTR dwParam1, DWORD_PTR dwParam2);
    void HandleSysEx(MIDIHDR* pHeader, DWORD_PTR dwParam2);
    void PrepareSysExBuffers();
    void UnprepareSysExBuffers();

    HMIDIIN m_hMidiIn = nullptr;
    std::atomic<bool> m_closing{false};
    MidiCallback m_callback;
    MidiSysExCallback m_sysExCallback;

    // MMCSS (Multimedia Class Scheduler Service)
    HANDLE m_mmcssHandle = nullptr;

    static constexpr int SYSEX_BUFFER_SIZE = 4096;
    static constexpr int SYSEX_BUFFER_COUNT = 4;
    MIDIHDR m_sysExHeaders[SYSEX_BUFFER_COUNT] = {};
    uint8_t m_sysExBuffers[SYSEX_BUFFER_COUNT][SYSEX_BUFFER_SIZE] = {};
};

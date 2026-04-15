#pragma once

#include <string>
#include <vector>
#include <unordered_set>
#include <mutex>
#include <windows.h>
#include <dsound.h>
#include "directmusic_defs.h"

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
    bool SendMidiMessage(uint8_t status, uint8_t data1, uint8_t data2, uint32_t channelGroup = 1);
    bool SendSysEx(const uint8_t* data, uint32_t length, uint32_t channelGroup = 1);
    void ResetAllState();

    const SynthConfig& GetActiveConfig() const { return m_activeConfig; }
    void Shutdown();

  private:
    bool DownloadGMInstruments();
    void SendGsReverbMacro(uint8_t type);
    void SendGsChorusMacro(uint8_t type);
    void SendGsDelayMacro(uint8_t type);
    void SendDefaultEffects();

    IDirectMusic8* m_pDirectMusic = nullptr;
    IDirectSound8* m_pDirectSound = nullptr;
    IDirectMusicPort* m_pPort = nullptr;
    IDirectMusicBuffer* m_pMusicBuffer = nullptr;

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
};

#pragma once
//
// DirectMusic COM interface definitions
// These are defined inline because DirectMusic headers were removed
// from the Windows SDK after the legacy DirectX SDK (June 2010).
// The COM objects (dmsynth.dll, dmime.dll, dmusic.dll) are still
// present and registered in Windows 10/11.
//

#include <windows.h>
#include <objbase.h>
#include <cguid.h>
#include <dsound.h>
#include <mmsystem.h>

// ============================================================
// GUIDs
// ============================================================

// CLSID_DirectMusic
// {636B9F10-0C7D-11D1-95B2-0020AFDC7421}
DEFINE_GUID(CLSID_DirectMusic, 0x636b9f10, 0x0c7d, 0x11d1, 0x95, 0xb2, 0x00, 0x20, 0xaf, 0xdc, 0x74, 0x21);

// IID_IDirectMusic8
// {2D3629F7-813D-4939-8508-F05C6B75FD97}
DEFINE_GUID(IID_IDirectMusic8, 0x2d3629f7, 0x813d, 0x4939, 0x85, 0x08, 0xf0, 0x5c, 0x6b, 0x75, 0xfd, 0x97);

// IID_IDirectMusicPort
// {08F2D8C9-37C2-11D2-B9F9-0000F875AC12}
DEFINE_GUID(IID_IDirectMusicPort, 0x08f2d8c9, 0x37c2, 0x11d2, 0xb9, 0xf9, 0x00, 0x00, 0xf8, 0x75, 0xac, 0x12);

// IID_IDirectMusicBuffer
// {D2AC2878-B39B-11D1-8704-00600893B1BD}
DEFINE_GUID(IID_IDirectMusicBuffer, 0xd2ac2878, 0xb39b, 0x11d1, 0x87, 0x04, 0x00, 0x60, 0x08, 0x93, 0xb1, 0xbd);

// GUID_DMUS_PROP_GS_Hardware_Emulation
// Not needed but keeping for reference

// CLSID_DirectMusicCollection
// {480FF4B0-28B2-11D1-BEF7-00C04FBF8FEF}
DEFINE_GUID(CLSID_DirectMusicCollection, 0x480ff4b0, 0x28b2, 0x11d1, 0xbe, 0xf7, 0x00, 0xc0, 0x4f, 0xbf, 0x8f, 0xef);

// GUID for the default Microsoft Software Synthesizer port
// {B7902FE9-FB0A-11D1-9994-00A0C9633418}
DEFINE_GUID(GUID_DMUS_SYNTH_DEFAULT, 0xb7902fe9, 0xfb0a, 0x11d1, 0x99, 0x94, 0x00, 0xa0, 0xc9, 0x63, 0x34, 0x18);

// CLSID_DirectMusicLoader
DEFINE_GUID(CLSID_DirectMusicLoader, 0xd2ac2892, 0xb39b, 0x11d1, 0x87, 0x04, 0x00, 0x60, 0x08, 0x93, 0xb1, 0xbd);

// IID_IDirectMusicLoader8
DEFINE_GUID(IID_IDirectMusicLoader8, 0x19e7c08c, 0x0a44, 0x4e6a, 0xa1, 0x16, 0x59, 0x5a, 0x7c, 0xd5, 0xde, 0x8c);

// IID_IDirectMusicCollection (and 8)
// {D2AC287C-B39B-11D1-8704-00600893B1BD}
DEFINE_GUID(IID_IDirectMusicCollection, 0xd2ac287c, 0xb39b, 0x11d1, 0x87, 0x04, 0x00, 0x60, 0x08, 0x93, 0xb1, 0xbd);
#define IID_IDirectMusicCollection8 IID_IDirectMusicCollection

// GUID_DefaultGMCollection
DEFINE_GUID(GUID_DefaultGMCollection, 0xf17e8673, 0xc3b4, 0x11d1, 0x87, 0x0b, 0x00, 0x60, 0x08, 0x93, 0xb1, 0xbd);

// IKsControl IID (from dmksctrl.h)
DEFINE_GUID(IID_IKsControl, 0x28f54685, 0x06fd, 0x11d2, 0xb2, 0x7a, 0x00, 0xa0, 0xc9, 0x22, 0x31, 0x96);

// DirectMusic Property GUIDs
DEFINE_GUID(GUID_DMUS_PROP_WriteLatency, 0x268a50e0, 0x5e46, 0x11d2, 0xaf, 0xa1, 0x00, 0xaa, 0x00, 0x24, 0xd8, 0xb6);
DEFINE_GUID(GUID_DMUS_PROP_WriteBufferLength, 0xfe442721, 0x18c3, 0x11d2, 0x81, 0x85, 0x00, 0x60, 0x08, 0x33, 0x16,
            0xc1);

// ============================================================
// Constants
// ============================================================

#define DMUS_PC_INPUTCLASS 0
#define DMUS_PC_OUTPUTCLASS 1

#define DMUS_PC_DLS 0x00000001
#define DMUS_PC_EXTERNAL 0x00000002
#define DMUS_PC_SOFTWARESYNTH 0x00000004
#define DMUS_PC_MEMORYMAPPED 0x00000008
#define DMUS_PC_SHARED 0x00000010
#define DMUS_PC_DLS2 0x00000200
#define DMUS_PC_GMINHARDWARE 0x00000040
#define DMUS_PC_GSINHARDWARE 0x00000080
#define DMUS_PC_XGINHARDWARE 0x00000100
#define DMUS_PC_DIRECTSOUND 0x00000080
#define DMUS_PC_SHAREABLE 0x00000100
#define DMUS_PC_SYSTEMMEMORY 0x7FFFFFFF

// DMUS_PORTPARAMS valid params flags
#define DMUS_PORTPARAMS_VOICES 0x00000001
#define DMUS_PORTPARAMS_CHANNELGROUPS 0x00000002
#define DMUS_PORTPARAMS_AUDIOCHANNELS 0x00000004
#define DMUS_PORTPARAMS_SAMPLERATE 0x00000008
#define DMUS_PORTPARAMS_EFFECTS 0x00000020
#define DMUS_PORTPARAMS_SHARE 0x00000040
#define DMUS_PORTPARAMS_FEATURES 0x00000080

#define DMUS_EFFECT_NONE 0x00000000
#define DMUS_EFFECT_REVERB 0x00000001
#define DMUS_EFFECT_CHORUS 0x00000002
#define DMUS_EFFECT_DELAY 0x00000004

// Event flags
#define DMUS_EVENT_STRUCTURED 0x00000001

// KSPROPERTY flags
#define KSPROPERTY_TYPE_GET 0x00000001
#define KSPROPERTY_TYPE_SET 0x00000002

// ============================================================
// Structures
// ============================================================

#pragma pack(push, 4)

typedef struct _DMUS_PORTCAPS
{
    DWORD dwSize;
    DWORD dwFlags;
    GUID guidPort;
    DWORD dwClass;
    DWORD dwType;
    DWORD dwMemorySize;
    DWORD dwMaxChannelGroups;
    DWORD dwMaxVoices;
    DWORD dwMaxAudioChannels;
    DWORD dwEffectFlags;
    WCHAR wszDescription[128];
} DMUS_PORTCAPS;

typedef struct _DMUS_PORTPARAMS8
{
    DWORD dwSize;
    DWORD dwValidParams;
    DWORD dwVoices;
    DWORD dwChannelGroups;
    DWORD dwAudioChannels;
    DWORD dwSampleRate;
    DWORD dwEffectFlags;
    DWORD fShare;
    DWORD dwFeatures;
} DMUS_PORTPARAMS8;

#define DMUS_OBJ_OBJECT 0x0001
#define DMUS_OBJ_CLASS 0x0002
#define DMUS_OBJ_NAME 0x0004
#define DMUS_OBJ_FILE 0x0008

typedef struct _DMUS_VERSION
{
    DWORD dwVersionMS;
    DWORD dwVersionLS;
} DMUS_VERSION;

typedef struct _DMUS_OBJECTDESC
{
    DWORD dwSize;
    DWORD dwValidData;
    GUID guidClass;
    GUID guidObject;
    GUID guidInstance;
    FILETIME ftDate;
    DMUS_VERSION vVersion;
    WCHAR wszName[64];
    WCHAR wszCategory[64];
    WCHAR wszFileName[260];
    IStream* pStream;
} DMUS_OBJECTDESC;

typedef struct _DMUS_BUFFERDESC
{
    DWORD dwSize;
    DWORD dwFlags;
    GUID guidBufferFormat;
    DWORD cbBuffer;
} DMUS_BUFFERDESC;

typedef struct _DMUS_NOTERANGE
{
    DWORD dwLowNote;
    DWORD dwHighNote;
} DMUS_NOTERANGE;

typedef struct
{
    GUID Set;
    ULONG Id;
    ULONG Flags;
} KSPROPERTY;

#pragma pack(pop)

// ============================================================
// Forward declarations
// ============================================================

struct IDirectMusicBuffer;
struct IDirectMusicPort;
struct IDirectMusic;
struct IDirectMusic8;
struct IReferenceClock;
struct IDirectMusicDownloadedInstrument;
struct IDirectMusicInstrument;

// ============================================================
// IReferenceClock - already defined in dsound.h (Windows SDK)
// We use it from there via the forward declaration in dsound.h.
// ============================================================

// ============================================================
// IDirectMusicBuffer
// ============================================================

#undef INTERFACE
#define INTERFACE IDirectMusicBuffer
DECLARE_INTERFACE_(IDirectMusicBuffer, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IDirectMusicBuffer
    STDMETHOD(Flush)(THIS) PURE;
    STDMETHOD(TotalTime)(THIS_ REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(PackStructured)(THIS_ REFERENCE_TIME rt, DWORD dwChannelGroup, DWORD dwChannelMessage) PURE;
    STDMETHOD(PackUnstructured)(THIS_ REFERENCE_TIME rt, DWORD dwChannelGroup, DWORD cb, BYTE * lpb) PURE;
    STDMETHOD(ResetReadPtr)(THIS) PURE;
    STDMETHOD(GetNextEvent)(THIS_ REFERENCE_TIME * prt, DWORD * pdwChannelGroup, DWORD * pdwLength, BYTE * *ppData)
        PURE;
    STDMETHOD(GetRawBufferPtr)(THIS_ BYTE * *ppData) PURE;
    STDMETHOD(GetStartTime)(THIS_ REFERENCE_TIME * prt) PURE;
    STDMETHOD(GetUsedBytes)(THIS_ DWORD * pcb) PURE;
    STDMETHOD(GetMaxBytes)(THIS_ DWORD * pcb) PURE;
    STDMETHOD(GetBufferFormat)(THIS_ GUID * pGuidFormat) PURE;
    STDMETHOD(SetStartTime)(THIS_ REFERENCE_TIME rt) PURE;
    STDMETHOD(SetUsedBytes)(THIS_ DWORD cb) PURE;
};

// ============================================================
// IDirectMusicPort
// ============================================================

#undef INTERFACE
#define INTERFACE IDirectMusicPort
DECLARE_INTERFACE_(IDirectMusicPort, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IDirectMusicPort
    STDMETHOD(PlayBuffer)(THIS_ IDirectMusicBuffer * pBuffer) PURE;
    STDMETHOD(SetReadNotificationHandle)(THIS_ HANDLE hEvent) PURE;
    STDMETHOD(Read)(THIS_ IDirectMusicBuffer * pBuffer) PURE;
    STDMETHOD(DownloadInstrument)(THIS_ void* pInstrument, void** ppDownloadedInstrument, void* pNoteRanges,
                                  DWORD dwNumNoteRanges) PURE;
    STDMETHOD(UnloadInstrument)(THIS_ void* pDownloadedInstrument) PURE;
    STDMETHOD(GetLatencyClock)(THIS_ IReferenceClock * *ppClock) PURE;
    STDMETHOD(GetRunningStats)(THIS_ void* pStats) PURE;
    STDMETHOD(Compact)(THIS) PURE;
    STDMETHOD(GetCaps)(THIS_ DMUS_PORTCAPS * pPortCaps) PURE;
    STDMETHOD(DeviceIoControl)(THIS_ DWORD dwIoControlCode, void* lpInBuffer, DWORD nInBufferSize, void* lpOutBuffer,
                               DWORD nOutBufferSize, DWORD* lpBytesReturned, OVERLAPPED* lpOverlapped) PURE;
    STDMETHOD(SetNumChannelGroups)(THIS_ DWORD dwChannelGroups) PURE;
    STDMETHOD(GetNumChannelGroups)(THIS_ DWORD * pdwChannelGroups) PURE;
    STDMETHOD(Activate)(THIS_ BOOL fActive) PURE;
    STDMETHOD(SetChannelPriority)(THIS_ DWORD dwChannelGroup, DWORD dwChannel, DWORD dwPriority) PURE;
    STDMETHOD(GetChannelPriority)(THIS_ DWORD dwChannelGroup, DWORD dwChannel, DWORD * pdwPriority) PURE;
    STDMETHOD(SetDirectSound)(THIS_ IDirectSound * pDirectSound, IDirectSoundBuffer * pDirectSoundBuffer) PURE;
    STDMETHOD(GetFormat)(THIS_ WAVEFORMATEX * pWaveFormatEx, DWORD * pdwWaveFormatExSize, DWORD * pdwBufferSize) PURE;
};

// ============================================================
// IDirectMusic / IDirectMusic8
// ============================================================

// IID_IDirectMusic
// {6536115A-7B2D-11D2-BA18-0000F875AC12}
DEFINE_GUID(IID_IDirectMusic, 0x6536115a, 0x7b2d, 0x11d2, 0xba, 0x18, 0x00, 0x00, 0xf8, 0x75, 0xac, 0x12);

#undef INTERFACE
#define INTERFACE IDirectMusic
DECLARE_INTERFACE_(IDirectMusic, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IDirectMusic
    STDMETHOD(EnumPort)(THIS_ DWORD dwIndex, DMUS_PORTCAPS * pPortCaps) PURE;
    STDMETHOD(CreateMusicBuffer)(THIS_ DMUS_BUFFERDESC * pBufferDesc, IDirectMusicBuffer * *ppBuffer,
                                 IUnknown * pUnkOuter) PURE;
    STDMETHOD(CreatePort)(THIS_ REFGUID rclsidPort, DMUS_PORTPARAMS8 * pPortParams, IDirectMusicPort * *ppPort,
                          IUnknown * pUnkOuter) PURE;
    STDMETHOD(EnumMasterClock)(THIS_ DWORD dwIndex, void* lpClockInfo) PURE;
    STDMETHOD(GetMasterClock)(THIS_ GUID * pguidClock, IReferenceClock * *ppReferenceClock) PURE;
    STDMETHOD(SetMasterClock)(THIS_ REFGUID rguidClock) PURE;
    STDMETHOD(Activate)(THIS_ BOOL fEnable) PURE;
    STDMETHOD(GetDefaultPort)(THIS_ GUID * pguidPort) PURE;
    STDMETHOD(SetDirectSound)(THIS_ IDirectSound * pDirectSound, HWND hWnd) PURE;
};

// ============================================================
// MUSIC_TIME type
// ============================================================
typedef long MUSIC_TIME;

#undef INTERFACE
#define INTERFACE IDirectMusic8
DECLARE_INTERFACE_(IDirectMusic8, IDirectMusic)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IDirectMusic
    STDMETHOD(EnumPort)(THIS_ DWORD dwIndex, DMUS_PORTCAPS * pPortCaps) PURE;
    STDMETHOD(CreateMusicBuffer)(THIS_ DMUS_BUFFERDESC * pBufferDesc, IDirectMusicBuffer * *ppBuffer,
                                 IUnknown * pUnkOuter) PURE;
    STDMETHOD(CreatePort)(THIS_ REFGUID rclsidPort, DMUS_PORTPARAMS8 * pPortParams, IDirectMusicPort * *ppPort,
                          IUnknown * pUnkOuter) PURE;
    STDMETHOD(EnumMasterClock)(THIS_ DWORD dwIndex, void* lpClockInfo) PURE;
    STDMETHOD(GetMasterClock)(THIS_ GUID * pguidClock, IReferenceClock * *ppReferenceClock) PURE;
    STDMETHOD(SetMasterClock)(THIS_ REFGUID rguidClock) PURE;
    STDMETHOD(Activate)(THIS_ BOOL fEnable) PURE;
    STDMETHOD(GetDefaultPort)(THIS_ GUID * pguidPort) PURE;
    STDMETHOD(SetDirectSound)(THIS_ IDirectSound * pDirectSound, HWND hWnd) PURE;

    // IDirectMusic8
    STDMETHOD(SetExternalMasterClock)(THIS_ IReferenceClock * pClock) PURE;
};

// ============================================================
// DLS Interfaces Stubs
// ============================================================

#undef INTERFACE
#define INTERFACE IDirectMusicInstrument
DECLARE_INTERFACE_(IDirectMusicInstrument, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(GetPatch)(THIS_ DWORD * pdwPatch) PURE;
    STDMETHOD(SetPatch)(THIS_ DWORD dwPatch) PURE;
};

#undef INTERFACE
#define INTERFACE IDirectMusicDownloadedInstrument
DECLARE_INTERFACE_(IDirectMusicDownloadedInstrument, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
};

#undef INTERFACE
#define INTERFACE IDirectMusicCollection
DECLARE_INTERFACE_(IDirectMusicCollection, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(GetInstrument)(THIS_ DWORD dwPatch, IDirectMusicInstrument * *ppInstrument) PURE;
    STDMETHOD(EnumInstrument)(THIS_ DWORD dwIndex, DWORD * pdwPatch, WCHAR * pwszName, DWORD dwNameLen) PURE;
};

// Using IDirectMusicCollection8 alias
typedef IDirectMusicCollection IDirectMusicCollection8;

#undef INTERFACE
#define INTERFACE IDirectMusicLoader8
DECLARE_INTERFACE_(IDirectMusicLoader8, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    STDMETHOD(GetObject)(THIS_ DMUS_OBJECTDESC * pDesc, REFIID riid, void** ppv) PURE;
    STDMETHOD(SetObject)(THIS_ DMUS_OBJECTDESC * pDesc) PURE;
    STDMETHOD(SetSearchDirectory)(THIS_ REFGUID rguidClass, WCHAR * pwzPath, BOOL fClear) PURE;
    STDMETHOD(ScanDirectory)(THIS_ REFGUID rguidClass, WCHAR * pwzFileExtension, WCHAR * pwzScanFileName) PURE;
    STDMETHOD(CacheObject)(THIS_ void* pObject) PURE;
    STDMETHOD(ReleaseObject)(THIS_ void* pObject) PURE;
    STDMETHOD(ClearCache)(THIS_ REFGUID rguidClass) PURE;
    STDMETHOD(EnableCache)(THIS_ REFGUID rguidClass, BOOL fEnable) PURE;
    STDMETHOD(EnumObject)(THIS_ REFGUID rguidClass, DWORD dwIndex, DMUS_OBJECTDESC * pDesc) PURE;
    STDMETHOD_(void, CollectGarbage)(THIS) PURE;
    STDMETHOD(ReleaseObjectByUnknown)(THIS_ IUnknown * pObject) PURE;
    STDMETHOD(LoadObjectFromFile)(THIS_ REFGUID rguidClassID, REFIID iidInterfaceID, WCHAR * pwzFilePath,
                                  void** ppObject) PURE;
};

// ============================================================
// IKsControl
// ============================================================
#undef INTERFACE
#define INTERFACE IKsControl
DECLARE_INTERFACE_(IKsControl, IUnknown)
{
    // IUnknown
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // IKsControl
    STDMETHOD(KsProperty)(THIS_ KSPROPERTY * Property, ULONG PropertyLength, void* PropertyData, ULONG DataLength,
                          ULONG* BytesReturned) PURE;
    STDMETHOD(KsMethod)(THIS_ void* Method, ULONG MethodLength, void* MethodData, ULONG DataLength,
                        ULONG* BytesReturned) PURE;
    STDMETHOD(KsEvent)(THIS_ void* Event, ULONG EventLength, void* EventData, ULONG DataLength, ULONG* BytesReturned)
        PURE;
};

// ============================================================
// IDirectMusicPerformance / IDirectMusicPerformance8
// Used for proper initialization of DirectMusic + DirectSound.
// On modern Windows, the low-level IDirectMusic::Activate path
// is broken; the Performance layer handles it correctly.
// ============================================================

// CLSID_DirectMusicPerformance
// {D2AC2881-B39B-11D1-8704-00600893B1BD}
DEFINE_GUID(CLSID_DirectMusicPerformance, 0xd2ac2881, 0xb39b, 0x11d1, 0x87, 0x04, 0x00, 0x60, 0x08, 0x93, 0xb1, 0xbd);

// IID_IDirectMusicPerformance
// {07D43D03-6523-11D2-871D-00600893B1BD}
DEFINE_GUID(IID_IDirectMusicPerformance, 0x07d43d03, 0x6523, 0x11d2, 0x87, 0x1d, 0x00, 0x60, 0x08, 0x93, 0xb1, 0xbd);

// IID_IDirectMusicPerformance8
// {679C4137-C929-11D1-A536-0000F875AC12}
DEFINE_GUID(IID_IDirectMusicPerformance8, 0x679c4137, 0xc929, 0x11d1, 0xa5, 0x36, 0x00, 0x00, 0xf8, 0x75, 0xac, 0x12);

// DMUS_AUDIOPARAMS for InitAudio
#define DMUS_AUDIOPARAMS_FEATURES 0x00000001
#define DMUS_AUDIOPARAMS_VOICES 0x00000002
#define DMUS_AUDIOPARAMS_SAMPLERATE 0x00000004
#define DMUS_AUDIOPARAMS_DEFAULTSYNTH 0x00000008

typedef struct _DMUS_AUDIOPARAMS
{
    DWORD dwSize;
    BOOL fInitNow;
    DWORD dwValidData;
    DWORD dwFeatures;
    DWORD dwVoices;
    DWORD dwSampleRate;
    CLSID clsidDefaultSynth;
} DMUS_AUDIOPARAMS;

// Audio path types for InitAudio
#define DMUS_APATH_SHARED_STEREOPLUSREVERB 1
#define DMUS_APATH_SHARED_STEREO 3
#define DMUS_APATH_DYNAMIC_MONO 7
#define DMUS_APATH_DYNAMIC_STEREO 8

// IDirectMusicPerformance - full vtable (41 methods after IUnknown)
#undef INTERFACE
#define INTERFACE IDirectMusicPerformance
DECLARE_INTERFACE_(IDirectMusicPerformance, IUnknown)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    STDMETHOD(Init)(THIS_ IDirectMusic * *ppDirectMusic, IDirectSound * pDirectSound, HWND hWnd) PURE;
    STDMETHOD(PlaySegment)(THIS_ void* pSegment, DWORD dwFlags, __int64 i64StartTime, void** ppSegmentState) PURE;
    STDMETHOD(Stop)(THIS_ void* pSegment, void* pSegmentState, MUSIC_TIME mtTime, DWORD dwFlags) PURE;
    STDMETHOD(GetSegmentState)(THIS_ void** ppSegmentState, MUSIC_TIME mtTime) PURE;
    STDMETHOD(SetPrepareTime)(THIS_ DWORD dwMilliSeconds) PURE;
    STDMETHOD(GetPrepareTime)(THIS_ DWORD * pdwMilliSeconds) PURE;
    STDMETHOD(SetBumperLength)(THIS_ DWORD dwMilliSeconds) PURE;
    STDMETHOD(GetBumperLength)(THIS_ DWORD * pdwMilliSeconds) PURE;
    STDMETHOD(SendPMsg)(THIS_ void* pPMSG) PURE;
    STDMETHOD(MusicToReferenceTime)(THIS_ MUSIC_TIME mtTime, REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(ReferenceToMusicTime)(THIS_ REFERENCE_TIME rtTime, MUSIC_TIME * pmtTime) PURE;
    STDMETHOD(IsPlaying)(THIS_ void* pSegment, void* pSegState) PURE;
    STDMETHOD(GetTime)(THIS_ REFERENCE_TIME * prtNow, MUSIC_TIME * pmtNow) PURE;
    STDMETHOD(AllocPMsg)(THIS_ ULONG cb, void** ppPMSG) PURE;
    STDMETHOD(FreePMsg)(THIS_ void* pPMSG) PURE;
    STDMETHOD(GetGraph)(THIS_ void** ppGraph) PURE;
    STDMETHOD(SetGraph)(THIS_ void* pGraph) PURE;
    STDMETHOD(SetNotificationHandle)(THIS_ HANDLE hNotification, REFERENCE_TIME rtMinimum) PURE;
    STDMETHOD(GetNotificationPMsg)(THIS_ void** ppNotificationPMsg) PURE;
    STDMETHOD(AddNotificationType)(THIS_ REFGUID rguidNotificationType) PURE;
    STDMETHOD(RemoveNotificationType)(THIS_ REFGUID rguidNotificationType) PURE;
    STDMETHOD(AddPort)(THIS_ IDirectMusicPort * pPort) PURE;
    STDMETHOD(RemovePort)(THIS_ IDirectMusicPort * pPort) PURE;
    STDMETHOD(AssignPChannelBlock)(THIS_ DWORD dwBlockNum, IDirectMusicPort * pPort, DWORD dwGroup) PURE;
    STDMETHOD(AssignPChannel)(THIS_ DWORD dwPChannel, IDirectMusicPort * pPort, DWORD dwGroup, DWORD dwMChannel) PURE;
    STDMETHOD(PChannelInfo)(THIS_ DWORD dwPChannel, IDirectMusicPort * *ppPort, DWORD * pdwGroup, DWORD * pdwMChannel)
        PURE;
    STDMETHOD(DownloadInstrument)(THIS_ void* pInst, DWORD dwPChannel, void** ppDownInst, void* pNoteRanges,
                                  DWORD dwNumNoteRanges, IDirectMusicPort** ppPort, DWORD* pdwGroup, DWORD* pdwMChannel)
        PURE;
    STDMETHOD(Invalidate)(THIS_ MUSIC_TIME mtTime, DWORD dwFlags) PURE;
    STDMETHOD(GetParam)(THIS_ REFGUID rguidType, DWORD dwGroupBits, DWORD dwIndex, MUSIC_TIME mtTime,
                        MUSIC_TIME * pmtNext, void* pParam) PURE;
    STDMETHOD(SetParam)(THIS_ REFGUID rguidType, DWORD dwGroupBits, DWORD dwIndex, MUSIC_TIME mtTime, void* pParam)
        PURE;
    STDMETHOD(GetGlobalParam)(THIS_ REFGUID rguidType, void* pParam, DWORD dwSize) PURE;
    STDMETHOD(SetGlobalParam)(THIS_ REFGUID rguidType, void* pParam, DWORD dwSize) PURE;
    STDMETHOD(GetLatencyTime)(THIS_ REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(GetQueueTime)(THIS_ REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(AdjustTime)(THIS_ REFERENCE_TIME rtAmount) PURE;
    STDMETHOD(CloseDown)(THIS) PURE;
    STDMETHOD(GetResolvedTime)(THIS_ REFERENCE_TIME rtTime, REFERENCE_TIME * prtResolved, DWORD dwTimeResolveFlags)
        PURE;
    STDMETHOD(MIDIToMusic)(THIS_ BYTE bMIDIValue, void* pChord, BYTE bPlayMode, BYTE bChordLevel, WORD* pwMusicValue)
        PURE;
    STDMETHOD(MusicToMIDI)(THIS_ WORD wMusicValue, void* pChord, BYTE bPlayMode, BYTE bChordLevel, BYTE* pbMIDIValue)
        PURE;
    STDMETHOD(TimeToRhythm)(THIS_ MUSIC_TIME mtTime, void* pTimeSig, WORD* pwMeasure, BYTE* pbBeat, BYTE* pbGrid,
                            short* pnOffset) PURE;
    STDMETHOD(RhythmToTime)(THIS_ WORD wMeasure, BYTE bBeat, BYTE bGrid, short nOffset, void* pTimeSig,
                            MUSIC_TIME* pmtTime) PURE;
};

// IDirectMusicPerformance8 extends IDirectMusicPerformance with InitAudio
#undef INTERFACE
#define INTERFACE IDirectMusicPerformance8
DECLARE_INTERFACE_(IDirectMusicPerformance8, IDirectMusicPerformance)
{
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppvObject) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    STDMETHOD(Init)(THIS_ IDirectMusic * *ppDirectMusic, IDirectSound * pDirectSound, HWND hWnd) PURE;
    STDMETHOD(PlaySegment)(THIS_ void* pSegment, DWORD dwFlags, __int64 i64StartTime, void** ppSegmentState) PURE;
    STDMETHOD(Stop)(THIS_ void* pSegment, void* pSegmentState, MUSIC_TIME mtTime, DWORD dwFlags) PURE;
    STDMETHOD(GetSegmentState)(THIS_ void** ppSegmentState, MUSIC_TIME mtTime) PURE;
    STDMETHOD(SetPrepareTime)(THIS_ DWORD dwMilliSeconds) PURE;
    STDMETHOD(GetPrepareTime)(THIS_ DWORD * pdwMilliSeconds) PURE;
    STDMETHOD(SetBumperLength)(THIS_ DWORD dwMilliSeconds) PURE;
    STDMETHOD(GetBumperLength)(THIS_ DWORD * pdwMilliSeconds) PURE;
    STDMETHOD(SendPMsg)(THIS_ void* pPMSG) PURE;
    STDMETHOD(MusicToReferenceTime)(THIS_ MUSIC_TIME mtTime, REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(ReferenceToMusicTime)(THIS_ REFERENCE_TIME rtTime, MUSIC_TIME * pmtTime) PURE;
    STDMETHOD(IsPlaying)(THIS_ void* pSegment, void* pSegState) PURE;
    STDMETHOD(GetTime)(THIS_ REFERENCE_TIME * prtNow, MUSIC_TIME * pmtNow) PURE;
    STDMETHOD(AllocPMsg)(THIS_ ULONG cb, void** ppPMSG) PURE;
    STDMETHOD(FreePMsg)(THIS_ void* pPMSG) PURE;
    STDMETHOD(GetGraph)(THIS_ void** ppGraph) PURE;
    STDMETHOD(SetGraph)(THIS_ void* pGraph) PURE;
    STDMETHOD(SetNotificationHandle)(THIS_ HANDLE hNotification, REFERENCE_TIME rtMinimum) PURE;
    STDMETHOD(GetNotificationPMsg)(THIS_ void** ppNotificationPMsg) PURE;
    STDMETHOD(AddNotificationType)(THIS_ REFGUID rguidNotificationType) PURE;
    STDMETHOD(RemoveNotificationType)(THIS_ REFGUID rguidNotificationType) PURE;
    STDMETHOD(AddPort)(THIS_ IDirectMusicPort * pPort) PURE;
    STDMETHOD(RemovePort)(THIS_ IDirectMusicPort * pPort) PURE;
    STDMETHOD(AssignPChannelBlock)(THIS_ DWORD dwBlockNum, IDirectMusicPort * pPort, DWORD dwGroup) PURE;
    STDMETHOD(AssignPChannel)(THIS_ DWORD dwPChannel, IDirectMusicPort * pPort, DWORD dwGroup, DWORD dwMChannel) PURE;
    STDMETHOD(PChannelInfo)(THIS_ DWORD dwPChannel, IDirectMusicPort * *ppPort, DWORD * pdwGroup, DWORD * pdwMChannel)
        PURE;
    STDMETHOD(DownloadInstrument)(THIS_ void* pInst, DWORD dwPChannel, void** ppDownInst, void* pNoteRanges,
                                  DWORD dwNumNoteRanges, IDirectMusicPort** ppPort, DWORD* pdwGroup, DWORD* pdwMChannel)
        PURE;
    STDMETHOD(Invalidate)(THIS_ MUSIC_TIME mtTime, DWORD dwFlags) PURE;
    STDMETHOD(GetParam)(THIS_ REFGUID rguidType, DWORD dwGroupBits, DWORD dwIndex, MUSIC_TIME mtTime,
                        MUSIC_TIME * pmtNext, void* pParam) PURE;
    STDMETHOD(SetParam)(THIS_ REFGUID rguidType, DWORD dwGroupBits, DWORD dwIndex, MUSIC_TIME mtTime, void* pParam)
        PURE;
    STDMETHOD(GetGlobalParam)(THIS_ REFGUID rguidType, void* pParam, DWORD dwSize) PURE;
    STDMETHOD(SetGlobalParam)(THIS_ REFGUID rguidType, void* pParam, DWORD dwSize) PURE;
    STDMETHOD(GetLatencyTime)(THIS_ REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(GetQueueTime)(THIS_ REFERENCE_TIME * prtTime) PURE;
    STDMETHOD(AdjustTime)(THIS_ REFERENCE_TIME rtAmount) PURE;
    STDMETHOD(CloseDown)(THIS) PURE;
    STDMETHOD(GetResolvedTime)(THIS_ REFERENCE_TIME rtTime, REFERENCE_TIME * prtResolved, DWORD dwTimeResolveFlags)
        PURE;
    STDMETHOD(MIDIToMusic)(THIS_ BYTE bMIDIValue, void* pChord, BYTE bPlayMode, BYTE bChordLevel, WORD* pwMusicValue)
        PURE;
    STDMETHOD(MusicToMIDI)(THIS_ WORD wMusicValue, void* pChord, BYTE bPlayMode, BYTE bChordLevel, BYTE* pbMIDIValue)
        PURE;
    STDMETHOD(TimeToRhythm)(THIS_ MUSIC_TIME mtTime, void* pTimeSig, WORD* pwMeasure, BYTE* pbBeat, BYTE* pbGrid,
                            short* pnOffset) PURE;
    STDMETHOD(RhythmToTime)(THIS_ WORD wMeasure, BYTE bBeat, BYTE bGrid, short nOffset, void* pTimeSig,
                            MUSIC_TIME* pmtTime) PURE;

    // IDirectMusicPerformance8 new methods
    STDMETHOD(InitAudio)(THIS_ IDirectMusic * *ppDirectMusic, IDirectSound * *ppDirectSound, HWND hWnd,
                         DWORD dwDefaultPathType, DWORD dwPChannelCount, DWORD dwFlags, DMUS_AUDIOPARAMS * pParams)
        PURE;
    STDMETHOD(PlaySegmentEx)(THIS_ IUnknown * pSource, WCHAR * pwzSegmentName, IUnknown * pTransition, DWORD dwFlags,
                             __int64 i64StartTime, void** ppSegmentState, IUnknown* pFrom, IUnknown* pAudioPath) PURE;
    STDMETHOD(StopEx)(THIS_ IUnknown * pObjectToStop, __int64 i64StopTime, DWORD dwFlags) PURE;
    STDMETHOD(ClonePMsg)(THIS_ void* pSourcePMSG, void** ppCopyPMSG) PURE;
    STDMETHOD(CreateAudioPath)(THIS_ IUnknown * pSourceConfig, BOOL fActivate, void** ppNewPath) PURE;
    STDMETHOD(CreateStandardAudioPath)(THIS_ DWORD dwType, DWORD dwPChannelCount, BOOL fActivate, void** ppNewPath)
        PURE;
    STDMETHOD(SetDefaultAudioPath)(THIS_ void* pAudioPath) PURE;
    STDMETHOD(GetDefaultAudioPath)(THIS_ void** ppAudioPath) PURE;
    STDMETHOD(GetParamEx)(THIS_ REFGUID rguidType, DWORD dwTrackID, DWORD dwGroupBits, DWORD dwIndex, MUSIC_TIME mtTime,
                          MUSIC_TIME * pmtNext, void* pParam) PURE;
};

#undef INTERFACE

#pragma once
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

#ifdef FVWRAP_EXPORTS
#define FVWRAP_API __declspec(dllexport)
#else
#define FVWRAP_API __declspec(dllimport)
#endif

typedef void* FVWRAP_HANDLE;

enum FVWRAP_ITEM_TYPE {
	FVWRAP_ITEM_NONE  = 0,
	FVWRAP_ITEM_AUDIO = 1,  // outAudio contains PCM bytes
	FVWRAP_ITEM_INDEX = 2,  // outValue = index
	FVWRAP_ITEM_DONE  = 3,  // utterance finished (no more audio/events for that commit)
	FVWRAP_ITEM_ERROR = 4   // outValue = error code
};

// Create wrapper. dataPath can be NULL (EngineFactory will try env/registry).
// speakerTavPath should usually be default.tav or Voices\Name.tav.
// language is numeric MM_TTSAPI::Language value (FVLanguage.h): e.g. 0x0409 for English.
FVWRAP_API FVWRAP_HANDLE __cdecl fvwrap_create(
	const char* dataPath,
	const char* speakerTavPath,
	int language,
	int sampleRate,
	int bitsPerSample
);

FVWRAP_API void __cdecl fvwrap_destroy(FVWRAP_HANDLE h);

// Stop current + queued speech, clear pending audio/events.
FVWRAP_API void __cdecl fvwrap_stop(FVWRAP_HANDLE h);

// Store-only settings (applied at safe boundaries between commits)
FVWRAP_API int __cdecl fvwrap_setRatePercent(FVWRAP_HANDLE h, int ratePercent);
FVWRAP_API int __cdecl fvwrap_setVolumePercent(FVWRAP_HANDLE h, int volPercent);
FVWRAP_API int __cdecl fvwrap_setPitchPercent(FVWRAP_HANDLE h, int pitchPercent);

// Build one utterance (text + index boundaries), then commit.
// The wrapper will internally inject BM_USER bookmarks for indexes.
FVWRAP_API void __cdecl fvwrap_begin(FVWRAP_HANDLE h);
FVWRAP_API void __cdecl fvwrap_addTextUtf8(FVWRAP_HANDLE h, const char* textUtf8);
FVWRAP_API void __cdecl fvwrap_addIndex(FVWRAP_HANDLE h, int index);
FVWRAP_API int  __cdecl fvwrap_commit(FVWRAP_HANDLE h, int repeatCount);

// Non-blocking: reads the next “stream item” in order.
// - If AUDIO: returns bytes copied into outAudio, sets outType=AUDIO.
// - If INDEX/DONE/ERROR: returns 0, sets outType + outValue.
// - If no item: outType=NONE, returns 0.
FVWRAP_API int __cdecl fvwrap_read(
	FVWRAP_HANDLE h,
	int* outType,
	int* outValue,
	uint8_t* outAudio,
	int outCap
);

#ifdef __cplusplus
}
#endif

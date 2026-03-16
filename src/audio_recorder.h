#ifndef AUDIO_RECORDER_H
#define AUDIO_RECORDER_H

#include <windows.h>

typedef struct AudioRecorderConfig {
	DWORD sample_rate;
	WORD channels;
	WORD bits_per_sample;
	SHORT voice_threshold;
	DWORD silence_timeout_ms;
	DWORD min_record_ms;
	DWORD max_record_ms;
} AudioRecorderConfig;

typedef struct AudioRuntimeStatus {
	BOOL is_recording;
	BOOL should_auto_stop;
	BOOL had_voice;
	DWORD peak_level;
	DWORD ms_since_voice;
	DWORD record_duration_ms;
	DWORD recorded_bytes;
} AudioRuntimeStatus;

BOOL audio_start_recording(const AudioRecorderConfig *config);
BOOL audio_stop_and_save(const wchar_t *wav_path);
void audio_abort(void);
BOOL audio_is_recording(void);
BOOL audio_get_runtime_status(AudioRuntimeStatus *out_status);

#endif

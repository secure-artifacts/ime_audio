#include "audio_recorder.h"

#include <mmsystem.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#pragma comment(lib, "winmm.lib")

#define RECORDER_BUFFER_COUNT 6
#define RECORDER_BUFFER_MS 100
#define VOICE_ACTIVATION_FRAMES 2

static CRITICAL_SECTION g_lock;
static BOOL g_lock_ready = FALSE;

static HWAVEIN g_wave_in = NULL;
static WAVEFORMATEX g_format;
static WAVEHDR g_headers[RECORDER_BUFFER_COUNT];
static BYTE *g_header_buffers[RECORDER_BUFFER_COUNT];

static BYTE *g_pcm_data = NULL;
static DWORD g_pcm_size = 0;
static DWORD g_pcm_capacity = 0;

static BOOL g_is_recording = FALSE;
static BOOL g_should_auto_stop = FALSE;
static BOOL g_has_voice = FALSE;
static DWORD g_voice_peak_hit_count = 0;
static DWORD g_last_peak = 0;

static ULONGLONG g_start_ms = 0;
static ULONGLONG g_last_voice_ms = 0;

static SHORT g_voice_threshold = 1400;
static DWORD g_silence_timeout_ms = 1500;
static DWORD g_min_record_ms = 900;
static DWORD g_max_record_ms = 30000;

static void ensure_lock(void) {
    if (!g_lock_ready) {
        InitializeCriticalSection(&g_lock);
        g_lock_ready = TRUE;
    }
}

static DWORD clamp_dword(DWORD value, DWORD min_value, DWORD max_value) {
    if (value < min_value) {
        return min_value;
    }
    if (value > max_value) {
        return max_value;
    }
    return value;
}

static BOOL append_pcm_data(const BYTE *data, DWORD size) {
    BYTE *new_buffer = NULL;
    DWORD required = 0;
    DWORD new_capacity = 0;

    if (!data || size == 0) {
        return TRUE;
    }

    if (size > MAXDWORD - g_pcm_size) {
        return FALSE;
    }
    required = g_pcm_size + size;

    if (required <= g_pcm_capacity) {
        memcpy(g_pcm_data + g_pcm_size, data, size);
        g_pcm_size = required;
        return TRUE;
    }

    new_capacity = g_pcm_capacity ? g_pcm_capacity : 65536;
    while (new_capacity < required) {
        if (new_capacity > MAXDWORD / 2) {
            new_capacity = required;
            break;
        }
        new_capacity *= 2;
    }

    new_buffer = (BYTE *)realloc(g_pcm_data, new_capacity);
    if (!new_buffer) {
        return FALSE;
    }

    g_pcm_data = new_buffer;
    g_pcm_capacity = new_capacity;
    memcpy(g_pcm_data + g_pcm_size, data, size);
    g_pcm_size = required;
    return TRUE;
}

static DWORD compute_peak_16bit(const BYTE *data, DWORD size) {
    const SHORT *samples = (const SHORT *)data;
    DWORD sample_count = size / sizeof(SHORT);
    DWORD peak = 0;
    DWORD i = 0;

    for (i = 0; i < sample_count; ++i) {
        int value = (int)samples[i];
        if (value < 0) {
            value = -value;
        }
        if ((DWORD)value > peak) {
            peak = (DWORD)value;
        }
    }

    return peak;
}

static void free_headers(void) {
    int i = 0;

    if (!g_wave_in) {
        for (i = 0; i < RECORDER_BUFFER_COUNT; ++i) {
            if (g_header_buffers[i]) {
                free(g_header_buffers[i]);
                g_header_buffers[i] = NULL;
            }
            ZeroMemory(&g_headers[i], sizeof(WAVEHDR));
        }
        return;
    }

    for (i = 0; i < RECORDER_BUFFER_COUNT; ++i) {
        if (g_headers[i].dwFlags & WHDR_PREPARED) {
            waveInUnprepareHeader(g_wave_in, &g_headers[i], sizeof(WAVEHDR));
        }
        if (g_header_buffers[i]) {
            free(g_header_buffers[i]);
            g_header_buffers[i] = NULL;
        }
        ZeroMemory(&g_headers[i], sizeof(WAVEHDR));
    }
}

static void close_wave_device(void) {
    if (!g_wave_in) {
        return;
    }

    waveInStop(g_wave_in);
    waveInReset(g_wave_in);
    free_headers();
    waveInClose(g_wave_in);
    g_wave_in = NULL;
}

static BOOL write_wav_file(const wchar_t *path) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    DWORD bytes_written = 0;
    DWORD riff_size = 0;
    DWORD fmt_chunk_size = 16;
    DWORD data_size = g_pcm_size;
    DWORD byte_rate = g_format.nAvgBytesPerSec;
    WORD block_align = g_format.nBlockAlign;

    if (!path) {
        return FALSE;
    }

    if (data_size > MAXDWORD - 36) {
        return FALSE;
    }
    riff_size = data_size + 36;

    file_handle = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file_handle == INVALID_HANDLE_VALUE) {
        return FALSE;
    }

    if (!WriteFile(file_handle, "RIFF", 4, &bytes_written, NULL) || bytes_written != 4 ||
        !WriteFile(file_handle, &riff_size, sizeof(riff_size), &bytes_written, NULL) || bytes_written != sizeof(riff_size) ||
        !WriteFile(file_handle, "WAVE", 4, &bytes_written, NULL) || bytes_written != 4 ||
        !WriteFile(file_handle, "fmt ", 4, &bytes_written, NULL) || bytes_written != 4 ||
        !WriteFile(file_handle, &fmt_chunk_size, sizeof(fmt_chunk_size), &bytes_written, NULL) || bytes_written != sizeof(fmt_chunk_size) ||
        !WriteFile(file_handle, &g_format.wFormatTag, sizeof(WORD), &bytes_written, NULL) || bytes_written != sizeof(WORD) ||
        !WriteFile(file_handle, &g_format.nChannels, sizeof(WORD), &bytes_written, NULL) || bytes_written != sizeof(WORD) ||
        !WriteFile(file_handle, &g_format.nSamplesPerSec, sizeof(DWORD), &bytes_written, NULL) || bytes_written != sizeof(DWORD) ||
        !WriteFile(file_handle, &byte_rate, sizeof(DWORD), &bytes_written, NULL) || bytes_written != sizeof(DWORD) ||
        !WriteFile(file_handle, &block_align, sizeof(WORD), &bytes_written, NULL) || bytes_written != sizeof(WORD) ||
        !WriteFile(file_handle, &g_format.wBitsPerSample, sizeof(WORD), &bytes_written, NULL) || bytes_written != sizeof(WORD) ||
        !WriteFile(file_handle, "data", 4, &bytes_written, NULL) || bytes_written != 4 ||
        !WriteFile(file_handle, &data_size, sizeof(data_size), &bytes_written, NULL) || bytes_written != sizeof(data_size)) {
        CloseHandle(file_handle);
        return FALSE;
    }

    if (data_size > 0) {
        if (!WriteFile(file_handle, g_pcm_data, data_size, &bytes_written, NULL) || bytes_written != data_size) {
            CloseHandle(file_handle);
            return FALSE;
        }
    }

    CloseHandle(file_handle);
    return TRUE;
}

static void reset_recording_data(void) {
    g_pcm_size = 0;
    g_should_auto_stop = FALSE;
    g_has_voice = FALSE;
    g_voice_peak_hit_count = 0;
    g_last_peak = 0;
    g_start_ms = 0;
    g_last_voice_ms = 0;
}

static void CALLBACK wave_in_proc(HWAVEIN hwi, UINT msg, DWORD_PTR instance, DWORD_PTR param1, DWORD_PTR param2) {
    WAVEHDR *header = NULL;
    BOOL requeue = FALSE;

    (void)hwi;
    (void)instance;
    (void)param2;

    if (msg != WIM_DATA) {
        return;
    }

    header = (WAVEHDR *)param1;
    if (!header) {
        return;
    }

    ensure_lock();
    EnterCriticalSection(&g_lock);

    if (g_is_recording && header->dwBytesRecorded > 0) {
        ULONGLONG now_ms = GetTickCount64();
        DWORD duration_ms = 0;
        DWORD ms_since_voice = 0;
        DWORD peak = compute_peak_16bit((const BYTE *)header->lpData, header->dwBytesRecorded);

        if (g_start_ms == 0) {
            g_start_ms = now_ms;
        }
        if (g_last_voice_ms == 0) {
            g_last_voice_ms = now_ms;
        }

        g_last_peak = peak;
        if (peak >= (DWORD)g_voice_threshold) {
            g_last_voice_ms = now_ms;
            if (g_voice_peak_hit_count < VOICE_ACTIVATION_FRAMES) {
                g_voice_peak_hit_count++;
            }
            if (g_voice_peak_hit_count >= VOICE_ACTIVATION_FRAMES) {
                g_has_voice = TRUE;
            }
        }

        append_pcm_data((const BYTE *)header->lpData, header->dwBytesRecorded);

        duration_ms = (DWORD)(now_ms - g_start_ms);
        ms_since_voice = (DWORD)(now_ms - g_last_voice_ms);

        if (g_has_voice && duration_ms >= g_min_record_ms && ms_since_voice >= g_silence_timeout_ms) {
            g_should_auto_stop = TRUE;
        }
        if (duration_ms >= g_max_record_ms) {
            g_should_auto_stop = TRUE;
        }
    }

    requeue = g_is_recording && g_wave_in != NULL;
    LeaveCriticalSection(&g_lock);

    if (requeue) {
        header->dwBytesRecorded = 0;
        waveInAddBuffer(g_wave_in, header, sizeof(WAVEHDR));
    }
}

BOOL audio_start_recording(const AudioRecorderConfig *config) {
    MMRESULT result;
    DWORD samples_per_buffer = 0;
    DWORD bytes_per_buffer = 0;
    int i = 0;

    ensure_lock();
    EnterCriticalSection(&g_lock);

    if (g_is_recording) {
        LeaveCriticalSection(&g_lock);
        return TRUE;
    }

    close_wave_device();
    reset_recording_data();

    g_format.wFormatTag = WAVE_FORMAT_PCM;
    g_format.nChannels = config && config->channels ? config->channels : 1;
    g_format.nSamplesPerSec = config && config->sample_rate ? config->sample_rate : 16000;
    g_format.wBitsPerSample = config && config->bits_per_sample ? config->bits_per_sample : 16;
    g_format.nBlockAlign = (WORD)((g_format.nChannels * g_format.wBitsPerSample) / 8);
    g_format.nAvgBytesPerSec = g_format.nSamplesPerSec * g_format.nBlockAlign;
    g_format.cbSize = 0;

    g_voice_threshold = config ? config->voice_threshold : 1400;
    g_silence_timeout_ms = config ? clamp_dword(config->silence_timeout_ms, 400, 6000) : 1500;
    g_min_record_ms = config ? clamp_dword(config->min_record_ms, 300, 5000) : 900;
    g_max_record_ms = config ? clamp_dword(config->max_record_ms, 3000, 120000) : 30000;

    if (g_voice_threshold < 150) {
        g_voice_threshold = 150;
    }
    if (g_voice_threshold > 6000) {
        g_voice_threshold = 6000;
    }

    result = waveInOpen(&g_wave_in, WAVE_MAPPER, &g_format, (DWORD_PTR)wave_in_proc, 0, CALLBACK_FUNCTION);
    if (result != MMSYSERR_NOERROR) {
        g_wave_in = NULL;
        LeaveCriticalSection(&g_lock);
        return FALSE;
    }

    samples_per_buffer = (g_format.nSamplesPerSec * RECORDER_BUFFER_MS) / 1000;
    if (samples_per_buffer == 0) {
        samples_per_buffer = g_format.nSamplesPerSec / 10;
        if (samples_per_buffer == 0) {
            samples_per_buffer = 1600;
        }
    }

    bytes_per_buffer = samples_per_buffer * g_format.nBlockAlign;
    if (bytes_per_buffer == 0) {
        bytes_per_buffer = 3200;
    }

    for (i = 0; i < RECORDER_BUFFER_COUNT; ++i) {
        ZeroMemory(&g_headers[i], sizeof(WAVEHDR));
        g_header_buffers[i] = (BYTE *)malloc(bytes_per_buffer);
        if (!g_header_buffers[i]) {
            close_wave_device();
            LeaveCriticalSection(&g_lock);
            return FALSE;
        }

        g_headers[i].lpData = (LPSTR)g_header_buffers[i];
        g_headers[i].dwBufferLength = bytes_per_buffer;

        result = waveInPrepareHeader(g_wave_in, &g_headers[i], sizeof(WAVEHDR));
        if (result != MMSYSERR_NOERROR) {
            close_wave_device();
            LeaveCriticalSection(&g_lock);
            return FALSE;
        }

        result = waveInAddBuffer(g_wave_in, &g_headers[i], sizeof(WAVEHDR));
        if (result != MMSYSERR_NOERROR) {
            close_wave_device();
            LeaveCriticalSection(&g_lock);
            return FALSE;
        }
    }

    result = waveInStart(g_wave_in);
    if (result != MMSYSERR_NOERROR) {
        close_wave_device();
        LeaveCriticalSection(&g_lock);
        return FALSE;
    }

    g_is_recording = TRUE;
    g_start_ms = GetTickCount64();
    g_last_voice_ms = g_start_ms;

    LeaveCriticalSection(&g_lock);
    return TRUE;
}

BOOL audio_stop_and_save(const wchar_t *wav_path) {
    BOOL was_recording = FALSE;
    BOOL write_ok = FALSE;

    if (!wav_path) {
        return FALSE;
    }

    ensure_lock();
    EnterCriticalSection(&g_lock);
    was_recording = g_is_recording;
    g_is_recording = FALSE;
    LeaveCriticalSection(&g_lock);

    if (!was_recording) {
        return FALSE;
    }

    close_wave_device();

    ensure_lock();
    EnterCriticalSection(&g_lock);
    write_ok = write_wav_file(wav_path) && g_pcm_size > 0;
    LeaveCriticalSection(&g_lock);

    return write_ok;
}

void audio_abort(void) {
    ensure_lock();
    EnterCriticalSection(&g_lock);
    g_is_recording = FALSE;
    LeaveCriticalSection(&g_lock);

    close_wave_device();

    ensure_lock();
    EnterCriticalSection(&g_lock);
    reset_recording_data();
    LeaveCriticalSection(&g_lock);
}

BOOL audio_is_recording(void) {
    BOOL is_recording = FALSE;

    ensure_lock();
    EnterCriticalSection(&g_lock);
    is_recording = g_is_recording;
    LeaveCriticalSection(&g_lock);

    return is_recording;
}

BOOL audio_get_runtime_status(AudioRuntimeStatus *out_status) {
    ULONGLONG now_ms;

    if (!out_status) {
        return FALSE;
    }

    ensure_lock();
    EnterCriticalSection(&g_lock);

    now_ms = GetTickCount64();
    out_status->is_recording = g_is_recording;
    out_status->should_auto_stop = g_should_auto_stop;
    out_status->had_voice = g_has_voice;
    out_status->peak_level = g_last_peak;
    out_status->recorded_bytes = g_pcm_size;

    if (g_start_ms > 0 && now_ms >= g_start_ms) {
        out_status->record_duration_ms = (DWORD)(now_ms - g_start_ms);
    } else {
        out_status->record_duration_ms = 0;
    }

    if (g_last_voice_ms > 0 && now_ms >= g_last_voice_ms) {
        out_status->ms_since_voice = (DWORD)(now_ms - g_last_voice_ms);
    } else {
        out_status->ms_since_voice = 0;
    }

    LeaveCriticalSection(&g_lock);
    return TRUE;
}

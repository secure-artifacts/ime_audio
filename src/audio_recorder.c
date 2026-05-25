#include "audio_recorder.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <objbase.h>
#include <mmdeviceapi.h>
#include <audioclient.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

#define VOICE_ACTIVATION_FRAMES 3

static CRITICAL_SECTION g_lock;
static BOOL g_lock_ready = FALSE;

// PCM buffer
static BYTE *g_pcm_data = NULL;
static DWORD g_pcm_size = 0;
static DWORD g_pcm_capacity = 0;

// Recording state
static BOOL g_is_recording = FALSE;
static BOOL g_should_auto_stop = FALSE;
static BOOL g_has_voice = FALSE;
static DWORD g_voice_peak_hit_count = 0;
static DWORD g_last_peak = 0;

static ULONGLONG g_start_ms = 0;
static ULONGLONG g_last_voice_ms = 0;
static double g_noise_floor = 0.0;

// Recording configuration
static WAVEFORMATEX g_format; // Output format (always 16kHz Mono 16-bit)
static SHORT g_voice_threshold = 1400;
static DWORD g_silence_timeout_ms = 1500;
static DWORD g_min_record_ms = 900;
static DWORD g_max_record_ms = 30000;

// WASAPI COM Interfaces
static IMMDevice *g_device = NULL;
static IAudioClient *g_audio_client = NULL;
static IAudioCaptureClient *g_capture_client = NULL;
static WAVEFORMATEX *g_mix_format = NULL;

// WASAPI Threading
static HANDLE g_record_thread = NULL;
static HANDLE g_record_stop_event = NULL;

// Fractional index state for resampler
static double g_resample_time = 0.0;

// Define WASAPI CLSIDs/IIDs locally to avoid linking issues
static const GUID CLSID_MMDeviceEnumerator_Local = {0xBCDE0395, 0xE52F, 0x467C, {0x8E, 0x3D, 0xC4, 0x57, 0x92, 0x91, 0x69, 0x2E}};
static const GUID IID_IMMDeviceEnumerator_Local = {0xA95664D2, 0x9614, 0x4F35, {0xA7, 0x46, 0xDE, 0x8D, 0xB6, 0x36, 0x17, 0xE6}};
static const GUID IID_IAudioClient_Local = {0x886D8EEB, 0x8CF2, 0x4446, {0x8D, 0x02, 0xCD, 0xBA, 0x1D, 0xBD, 0x71, 0x05}};
static const GUID IID_IAudioCaptureClient_Local = {0xC8ADBD64, 0xE71E, 0x48A0, {0xA4, 0xDE, 0x18, 0x5C, 0x39, 0x5C, 0xD3, 0x17}};
static const GUID KSDATAFORMAT_SUBTYPE_IEEE_FLOAT_Local = {0x00000003, 0x0000, 0x0010, {0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71}};
static const GUID KSDATAFORMAT_SUBTYPE_PCM_Local = {0x00000001, 0x0000, 0x0010, {0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71}};

static void ensure_lock(void) {
    if (!g_lock_ready) {
        InitializeCriticalSection(&g_lock);
        g_lock_ready = TRUE;
    }
}

static DWORD clamp_dword(DWORD value, DWORD min_value, DWORD max_value) {
    if (value < min_value) return min_value;
    if (value > max_value) return max_value;
    return value;
}

static DWORD clamp_setting(DWORD value, DWORD min_value, DWORD max_value) {
    return clamp_dword(value, min_value, max_value);
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

static float get_float_sample(const BYTE *pFrameData, int channel, const WAVEFORMATEX *format) {
    int bytes_per_sample = format->wBitsPerSample / 8;
    int offset = channel * bytes_per_sample;
    const BYTE *pSample = pFrameData + offset;
    
    if (format->wFormatTag == WAVE_FORMAT_PCM) {
        if (format->wBitsPerSample == 16) {
            short val = *(const short*)pSample;
            return (float)val / 32768.0f;
        } else if (format->wBitsPerSample == 24) {
            int val = (pSample[0] << 8) | (pSample[1] << 16) | (pSample[2] << 24);
            val >>= 8;
            return (float)val / 8388608.0f;
        } else if (format->wBitsPerSample == 32) {
            int val = *(const int*)pSample;
            return (float)val / 2147483648.0f;
        }
    } else if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        const WAVEFORMATEXTENSIBLE *pExt = (const WAVEFORMATEXTENSIBLE*)format;
        if (memcmp(&pExt->SubFormat, &KSDATAFORMAT_SUBTYPE_IEEE_FLOAT_Local, sizeof(GUID)) == 0) {
            return *(const float*)pSample;
        } else if (memcmp(&pExt->SubFormat, &KSDATAFORMAT_SUBTYPE_PCM_Local, sizeof(GUID)) == 0) {
            if (format->wBitsPerSample == 16) {
                short val = *(const short*)pSample;
                return (float)val / 32768.0f;
            } else if (format->wBitsPerSample == 24) {
                int val = (pSample[0] << 8) | (pSample[1] << 16) | (pSample[2] << 24);
                val >>= 8;
                return (float)val / 8388608.0f;
            } else if (format->wBitsPerSample == 32) {
                int val = *(const int*)pSample;
                return (float)val / 2147483648.0f;
            }
        }
    }
    return 0.0f;
}

static void close_wasapi_device(void) {
    if (g_record_thread) {
        if (g_record_stop_event) {
            SetEvent(g_record_stop_event);
        }
        g_is_recording = FALSE;
        WaitForSingleObject(g_record_thread, 2000);
        CloseHandle(g_record_thread);
        g_record_thread = NULL;
    }

    if (g_record_stop_event) {
        CloseHandle(g_record_stop_event);
        g_record_stop_event = NULL;
    }

    if (g_capture_client) {
        g_capture_client->lpVtbl->Release(g_capture_client);
        g_capture_client = NULL;
    }

    if (g_audio_client) {
        g_audio_client->lpVtbl->Stop(g_audio_client);
        g_audio_client->lpVtbl->Release(g_audio_client);
        g_audio_client = NULL;
    }

    if (g_device) {
        g_device->lpVtbl->Release(g_device);
        g_device = NULL;
    }

    if (g_mix_format) {
        CoTaskMemFree(g_mix_format);
        g_mix_format = NULL;
    }
}

static void reset_recording_data(void) {
    g_pcm_size = 0;
    g_should_auto_stop = FALSE;
    g_has_voice = FALSE;
    g_voice_peak_hit_count = 0;
    g_last_peak = 0;
    g_start_ms = 0;
    g_last_voice_ms = 0;
    g_noise_floor = 0.0;
}

static DWORD WINAPI wasapi_record_thread_proc(LPVOID param) {
    (void)param;
    
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        return 0;
    }
    
    double step = (double)g_mix_format->nSamplesPerSec / 16000.0;
    
    while (1) {
        DWORD wait_result = WaitForSingleObject(g_record_stop_event, 30);
        
        ensure_lock();
        EnterCriticalSection(&g_lock);
        BOOL should_exit = !g_is_recording;
        LeaveCriticalSection(&g_lock);
        
        if (should_exit || wait_result == WAIT_OBJECT_0) {
            break;
        }
        
        UINT32 packetSize = 0;
        hr = g_capture_client->lpVtbl->GetNextPacketSize(g_capture_client, &packetSize);
        if (FAILED(hr)) {
            break;
        }
        
        while (packetSize > 0) {
            BYTE *pData = NULL;
            UINT32 numFramesToRead = 0;
            DWORD flags = 0;
            
            hr = g_capture_client->lpVtbl->GetBuffer(g_capture_client, &pData, &numFramesToRead, &flags, NULL, NULL);
            if (SUCCEEDED(hr)) {
                if (numFramesToRead > 0) {
                    int max_out_frames = (int)((double)numFramesToRead / step) + 2;
                    short *out_samples = (short*)malloc(max_out_frames * sizeof(short));
                    int out_idx = 0;
                    
                    if (out_samples) {
                        while (g_resample_time < (double)numFramesToRead) {
                            int idx = (int)(g_resample_time + 0.5);
                            if (idx >= (int)numFramesToRead) {
                                idx = (int)numFramesToRead - 1;
                            }
                            if (idx < 0) {
                                idx = 0;
                            }
                            
                            float mono = 0.0f;
                            int channels = g_mix_format->nChannels;
                            const BYTE *pFrame = pData + idx * g_mix_format->nBlockAlign;
                            
                            if (flags & AUDCLNT_BUFFERFLAGS_SILENT) {
                                mono = 0.0f;
                            } else {
                                for (int c = 0; c < channels; c++) {
                                    mono += get_float_sample(pFrame, c, g_mix_format);
                                }
                                mono /= (float)channels;
                            }
                            
                            if (mono > 1.0f) mono = 1.0f;
                            if (mono < -1.0f) mono = -1.0f;
                            
                            out_samples[out_idx++] = (short)(mono * 32767.0f);
                            
                            g_resample_time += step;
                        }
                        g_resample_time -= (double)numFramesToRead;
                        
                        if (out_idx > 0) {
                            EnterCriticalSection(&g_lock);
                            if (g_is_recording) {
                                ULONGLONG now_ms = GetTickCount64();
                                DWORD duration_ms = 0;
                                DWORD ms_since_voice = 0;
                                
                                DWORD peak = 0;
                                for (int s = 0; s < out_idx; s++) {
                                    int val = out_samples[s];
                                    if (val < 0) val = -val;
                                    if ((DWORD)val > peak) peak = (DWORD)val;
                                }
                                
                                if (g_start_ms == 0) {
                                    g_start_ms = now_ms;
                                    g_noise_floor = (double)peak;
                                }
                                if (g_last_voice_ms == 0) {
                                    g_last_voice_ms = now_ms;
                                }
                                
                                if ((double)peak < g_noise_floor) {
                                    g_noise_floor = g_noise_floor * 0.2 + (double)peak * 0.8;
                                } else {
                                    g_noise_floor = g_noise_floor * 0.995 + (double)peak * 0.005;
                                }
                                
                                double capped_noise = g_noise_floor;
                                if (capped_noise > 3000.0) {
                                    capped_noise = 3000.0;
                                }
                                
                                DWORD adaptive_threshold = (DWORD)g_voice_threshold + (DWORD)capped_noise;
                                g_last_peak = peak;
                                
                                if (peak >= adaptive_threshold) {
                                    g_last_voice_ms = now_ms;
                                    if (g_voice_peak_hit_count < VOICE_ACTIVATION_FRAMES) {
                                        g_voice_peak_hit_count++;
                                    }
                                    if (g_voice_peak_hit_count >= VOICE_ACTIVATION_FRAMES) {
                                        g_has_voice = TRUE;
                                    }
                                } else {
                                    if (!g_has_voice && (now_ms - g_last_voice_ms > 400)) {
                                        g_voice_peak_hit_count = 0;
                                    }
                                }
                                
                                append_pcm_data((const BYTE*)out_samples, out_idx * sizeof(short));
                                
                                duration_ms = (DWORD)(now_ms - g_start_ms);
                                ms_since_voice = (DWORD)(now_ms - g_last_voice_ms);
                                
                                if (g_has_voice && duration_ms >= g_min_record_ms && ms_since_voice >= g_silence_timeout_ms) {
                                    g_should_auto_stop = TRUE;
                                }
                                if (duration_ms >= g_max_record_ms) {
                                    g_should_auto_stop = TRUE;
                                }
                            }
                            LeaveCriticalSection(&g_lock);
                        }
                        free(out_samples);
                    }
                }
                g_capture_client->lpVtbl->ReleaseBuffer(g_capture_client, numFramesToRead);
            }
            
            hr = g_capture_client->lpVtbl->GetNextPacketSize(g_capture_client, &packetSize);
            if (FAILED(hr)) {
                break;
            }
        }
    }
    
    CoUninitialize();
    return 0;
}

BOOL audio_start_recording(const AudioRecorderConfig *config) {
    HRESULT hr;
    IMMDeviceEnumerator *pEnumerator = NULL;
    IMMDeviceCollection *pCollection = NULL;
    IMMDevice *pTargetDevice = NULL;
    UINT count = 0;
    BOOL found = FALSE;

    ensure_lock();
    EnterCriticalSection(&g_lock);

    if (g_is_recording) {
        LeaveCriticalSection(&g_lock);
        return TRUE;
    }

    close_wasapi_device();
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

    if (g_voice_threshold < 150) g_voice_threshold = 150;
    if (g_voice_threshold > 6000) g_voice_threshold = 6000;

    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
        LeaveCriticalSection(&g_lock);
        return FALSE;
    }

    hr = CoCreateInstance(&CLSID_MMDeviceEnumerator_Local, NULL, CLSCTX_ALL, &IID_IMMDeviceEnumerator_Local, (void**)&pEnumerator);
    if (FAILED(hr)) {
        goto cleanup;
    }

    if (config && config->device_name[0] != L'\0' && wcscmp(config->device_name, L"系统默认录音设备") != 0) {
        hr = pEnumerator->lpVtbl->EnumAudioEndpoints(pEnumerator, eCapture, DEVICE_STATE_ACTIVE, &pCollection);
        if (SUCCEEDED(hr)) {
            pCollection->lpVtbl->GetCount(pCollection, &count);
            for (UINT i = 0; i < count; i++) {
                IMMDevice *pDev = NULL;
                hr = pCollection->lpVtbl->Item(pCollection, i, &pDev);
                if (SUCCEEDED(hr)) {
                    IPropertyStore *pProps = NULL;
                    hr = pDev->lpVtbl->OpenPropertyStore(pDev, STGM_READ, &pProps);
                    if (SUCCEEDED(hr)) {
                        PROPVARIANT varName;
                        PropVariantInit(&varName);
                        PROPERTYKEY key = {{0xa45c254e, 0xdf1c, 0x4efd, {0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0}}, 14};
                        hr = pProps->lpVtbl->GetValue(pProps, &key, &varName);
                        if (SUCCEEDED(hr)) {
                            if (wcsncmp(varName.pwszVal, config->device_name, wcslen(config->device_name)) == 0 ||
                                wcsncmp(config->device_name, varName.pwszVal, wcslen(varName.pwszVal)) == 0) {
                                pTargetDevice = pDev;
                                pTargetDevice->lpVtbl->AddRef(pTargetDevice);
                                found = TRUE;
                            }
                            PropVariantClear(&varName);
                        }
                        pProps->lpVtbl->Release(pProps);
                    }
                    pDev->lpVtbl->Release(pDev);
                }
                if (found) break;
            }
            pCollection->lpVtbl->Release(pCollection);
        }
    }

    if (!pTargetDevice) {
        hr = pEnumerator->lpVtbl->GetDefaultAudioEndpoint(pEnumerator, eCapture, eConsole, &pTargetDevice);
        if (FAILED(hr)) {
            goto cleanup;
        }
    }

    g_device = pTargetDevice;

    hr = g_device->lpVtbl->Activate(g_device, &IID_IAudioClient_Local, CLSCTX_ALL, NULL, (void**)&g_audio_client);
    if (FAILED(hr)) {
        goto cleanup;
    }

    hr = g_audio_client->lpVtbl->GetMixFormat(g_audio_client, &g_mix_format);
    if (FAILED(hr)) {
        goto cleanup;
    }

    hr = g_audio_client->lpVtbl->Initialize(g_audio_client, AUDCLNT_SHAREMODE_SHARED, 0, 10000000, 0, g_mix_format, NULL);
    if (FAILED(hr)) {
        goto cleanup;
    }

    hr = g_audio_client->lpVtbl->GetService(g_audio_client, &IID_IAudioCaptureClient_Local, (void**)&g_capture_client);
    if (FAILED(hr)) {
        goto cleanup;
    }

    hr = g_audio_client->lpVtbl->Start(g_audio_client);
    if (FAILED(hr)) {
        goto cleanup;
    }

    g_resample_time = 0.0;
    g_record_stop_event = CreateEventW(NULL, TRUE, FALSE, NULL);
    if (!g_record_stop_event) {
        goto cleanup;
    }

    g_is_recording = TRUE;
    g_start_ms = GetTickCount64();
    g_last_voice_ms = g_start_ms;

    g_record_thread = CreateThread(NULL, 0, wasapi_record_thread_proc, NULL, 0, NULL);
    if (!g_record_thread) {
        g_is_recording = FALSE;
        goto cleanup;
    }

    if (pEnumerator) pEnumerator->lpVtbl->Release(pEnumerator);
    CoUninitialize();
    LeaveCriticalSection(&g_lock);
    return TRUE;

cleanup:
    close_wasapi_device();
    if (pEnumerator) pEnumerator->lpVtbl->Release(pEnumerator);
    CoUninitialize();
    LeaveCriticalSection(&g_lock);
    return FALSE;
}

static BOOL write_wav_file(const wchar_t *path) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    DWORD bytes_written = 0;
    DWORD riff_size = 0;
    DWORD fmt_chunk_size = 16;
    DWORD data_size = g_pcm_size;
    DWORD byte_rate = g_format.nAvgBytesPerSec;
    WORD block_align = g_format.nBlockAlign;

    if (g_has_voice && g_start_ms > 0 && g_last_voice_ms >= g_start_ms) {
        ULONGLONG speech_ms = (g_last_voice_ms - g_start_ms) + 400; // 400ms margin
        DWORD bytes_per_ms = g_format.nAvgBytesPerSec / 1000;
        if (bytes_per_ms > 0) {
            DWORD speech_bytes = (DWORD)speech_ms * bytes_per_ms;
            speech_bytes = (speech_bytes / g_format.nBlockAlign) * g_format.nBlockAlign;
            if (speech_bytes < data_size) {
                data_size = speech_bytes;
            }
        }
    }

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

    close_wasapi_device();

    ensure_lock();
    EnterCriticalSection(&g_lock);
    write_ok = write_wav_file(wav_path) && g_pcm_size > 0;
    LeaveCriticalSection(&g_lock);

    return write_ok;
}

BOOL audio_save_chunk_and_continue(const wchar_t *wav_path) {
    BYTE *temp_pcm = NULL;
    DWORD temp_size = 0;
    BOOL write_ok = FALSE;

    if (!wav_path) {
        return FALSE;
    }

    ensure_lock();
    EnterCriticalSection(&g_lock);

    if (!g_is_recording || g_pcm_size == 0) {
        LeaveCriticalSection(&g_lock);
        return FALSE;
    }

    temp_size = g_pcm_size;
    if (g_has_voice && g_start_ms > 0 && g_last_voice_ms >= g_start_ms) {
        ULONGLONG speech_ms = (g_last_voice_ms - g_start_ms) + 400; // 400ms margin
        DWORD bytes_per_ms = g_format.nAvgBytesPerSec / 1000;
        if (bytes_per_ms > 0) {
            DWORD speech_bytes = (DWORD)speech_ms * bytes_per_ms;
            speech_bytes = (speech_bytes / g_format.nBlockAlign) * g_format.nBlockAlign;
            if (speech_bytes < temp_size) {
                temp_size = speech_bytes;
            }
        }
    }

    temp_pcm = (BYTE *)malloc(temp_size);
    if (temp_pcm) {
        memcpy(temp_pcm, g_pcm_data, temp_size);
    }

    g_pcm_size = 0;
    g_should_auto_stop = FALSE;
    g_has_voice = FALSE;
    g_voice_peak_hit_count = 0;
    g_start_ms = GetTickCount64();
    g_last_voice_ms = g_start_ms;

    LeaveCriticalSection(&g_lock);

    if (!temp_pcm) {
        return FALSE;
    }

    {
        HANDLE file_handle = INVALID_HANDLE_VALUE;
        DWORD bytes_written = 0;
        DWORD riff_size = temp_size + 36;
        DWORD fmt_chunk_size = 16;
        DWORD byte_rate = g_format.nAvgBytesPerSec;
        WORD block_align = g_format.nBlockAlign;

        file_handle = CreateFileW(wav_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (file_handle != INVALID_HANDLE_VALUE) {
            WriteFile(file_handle, "RIFF", 4, &bytes_written, NULL);
            WriteFile(file_handle, &riff_size, 4, &bytes_written, NULL);
            WriteFile(file_handle, "WAVE", 4, &bytes_written, NULL);
            WriteFile(file_handle, "fmt ", 4, &bytes_written, NULL);
            WriteFile(file_handle, &fmt_chunk_size, 4, &bytes_written, NULL);
            WriteFile(file_handle, &g_format.wFormatTag, 2, &bytes_written, NULL);
            WriteFile(file_handle, &g_format.nChannels, 2, &bytes_written, NULL);
            WriteFile(file_handle, &g_format.nSamplesPerSec, 4, &bytes_written, NULL);
            WriteFile(file_handle, &byte_rate, 4, &bytes_written, NULL);
            WriteFile(file_handle, &block_align, 2, &bytes_written, NULL);
            WriteFile(file_handle, &g_format.wBitsPerSample, 2, &bytes_written, NULL);
            WriteFile(file_handle, "data", 4, &bytes_written, NULL);
            WriteFile(file_handle, &temp_size, 4, &bytes_written, NULL);
            WriteFile(file_handle, temp_pcm, temp_size, &bytes_written, NULL);
            CloseHandle(file_handle);
            write_ok = (bytes_written == temp_size);
        }
    }

    free(temp_pcm);
    return write_ok;
}

void audio_abort(void) {
    ensure_lock();
    EnterCriticalSection(&g_lock);
    g_is_recording = FALSE;
    LeaveCriticalSection(&g_lock);

    close_wasapi_device();

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

#include "asr_backend.h"

#include <ctype.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <winhttp.h>

#define SHERPA_OUTPUT_MAX (256 * 1024)

static char *json_escape(const char *input);

static char *wide_to_utf8_alloc(const wchar_t *wide_text) {
    int needed = 0;
    char *utf8 = NULL;

    if (!wide_text) {
        return NULL;
    }

    needed = WideCharToMultiByte(CP_UTF8, 0, wide_text, -1, NULL, 0, NULL, NULL);
    if (needed <= 0) {
        return NULL;
    }

    utf8 = (char *)malloc((size_t)needed);
    if (!utf8) {
        return NULL;
    }

    if (WideCharToMultiByte(CP_UTF8, 0, wide_text, -1, utf8, needed, NULL, NULL) <= 0) {
        free(utf8);
        return NULL;
    }

    return utf8;
}

static char *base64_encode(const unsigned char *data, size_t len, size_t *out_len) {
    size_t encoded_len = 4 * ((len + 2) / 3);
    char *out = (char *)malloc(encoded_len + 1);
    if (!out) return NULL;

    static const char base64_chars[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    
    size_t i = 0, j = 0;
    for (i = 0; i < len;) {
        uint32_t octet_a = i < len ? data[i++] : 0;
        uint32_t octet_b = i < len ? data[i++] : 0;
        uint32_t octet_c = i < len ? data[i++] : 0;

        uint32_t triple = (octet_a << 16) + (octet_b << 8) + octet_c;

        out[j++] = base64_chars[(triple >> 18) & 0x3F];
        out[j++] = base64_chars[(triple >> 12) & 0x3F];
        out[j++] = base64_chars[(triple >> 6) & 0x3F];
        out[j++] = base64_chars[(triple) & 0x3F];
    }
    
    for (int k = 0; k < (int)((3 - len % 3) % 3); k++) {
        out[encoded_len - 1 - k] = '=';
    }
    out[encoded_len] = '\0';

    if (out_len) {
        *out_len = encoded_len;
    }
    return out;
}

static char *dup_text_limited(const char *text, size_t max_len) {
    size_t len = 0;
    char *copy = NULL;

    if (!text) {
        return NULL;
    }

    len = strlen(text);
    if (len > max_len) {
        len = max_len;
    }

    copy = (char *)malloc(len + 1);
    if (!copy) {
        return NULL;
    }

    memcpy(copy, text, len);
    copy[len] = '\0';
    return copy;
}

static void set_error_text(char **out_error_utf8, const char *message) {
    if (!out_error_utf8 || !message) {
        return;
    }

    free(*out_error_utf8);
    *out_error_utf8 = dup_text_limited(message, 1800);
}

static void trim_ascii_whitespace(char *text) {
    char *start = text;
    char *end = NULL;

    if (!text) {
        return;
    }

    while (*start == ' ' || *start == '\t' || *start == '\r' || *start == '\n') {
        start++;
    }

    if (start != text) {
        memmove(text, start, strlen(start) + 1);
    }

    end = text + strlen(text);
    while (end > text && (end[-1] == ' ' || end[-1] == '\t' || end[-1] == '\r' || end[-1] == '\n')) {
        end--;
    }

    *end = '\0';
}

static BOOL is_empty_wide(const wchar_t *text) {
    return !text || text[0] == L'\0';
}

static BOOL append_chunk(char **target, size_t *target_len, const char *data, size_t data_len) {
    char *new_buffer = NULL;

    if (!target || !target_len || !data || data_len == 0) {
        return TRUE;
    }

    if (*target_len > SHERPA_OUTPUT_MAX || data_len > SHERPA_OUTPUT_MAX || *target_len + data_len > SHERPA_OUTPUT_MAX) {
        return FALSE;
    }

    new_buffer = (char *)realloc(*target, *target_len + data_len + 1);
    if (!new_buffer) {
        return FALSE;
    }

    *target = new_buffer;
    memcpy(*target + *target_len, data, data_len);
    *target_len += data_len;
    (*target)[*target_len] = '\0';
    return TRUE;
}

static void extract_directory(const wchar_t *path, wchar_t *out_dir, size_t out_dir_len) {
    size_t len = 0;

    if (!path || !out_dir || out_dir_len == 0) {
        return;
    }

    out_dir[0] = L'\0';
    len = wcslen(path);

    if (len == 0 || len >= out_dir_len) {
        return;
    }

    wcscpy_s(out_dir, out_dir_len, path);
    while (len > 0) {
        if (out_dir[len - 1] == L'\\' || out_dir[len - 1] == L'/') {
            out_dir[len - 1] = L'\0';
            return;
        }
        len--;
    }

    out_dir[0] = L'\0';
}

static BOOL is_separator_only_line(const char *line) {
    const unsigned char *p = (const unsigned char *)line;
    size_t visible_chars = 0;

    if (!line || line[0] == '\0') {
        return FALSE;
    }

    while (*p) {
        unsigned char c = *p;

        if (c == ' ' || c == '\t' || c == '\r' || c == '\n') {
            p++;
            continue;
        }

        visible_chars++;
        if (c != '-' && c != '=' && c != '_' && c != '*' && c != '.' && c != ':' &&
            c != '~' && c != '|' && c != '+' && c != '#') {
            return FALSE;
        }

        p++;
    }

    return visible_chars >= 3;
}

static BOOL seems_transcription_text(const char *line) {
    const unsigned char *p = (const unsigned char *)line;

    if (!line || line[0] == '\0') {
        return FALSE;
    }

    if (is_separator_only_line(line)) {
        return FALSE;
    }

    while (*p) {
        unsigned char c = *p;
        if (c >= 0x80 || isalnum(c)) {
            return TRUE;
        }
        p++;
    }

    return FALSE;
}

static BOOL is_noisy_status_line(const char *line) {
    if (!line || line[0] == '\0') {
        return TRUE;
    }

    if (is_separator_only_line(line)) {
        return TRUE;
    }

    if (strstr(line, "Loading model") != NULL ||
        strstr(line, "Please check your config") != NULL ||
        strstr(line, "sherpa-onnx") != NULL ||
        strstr(line, "OfflineRecognizerConfig(") != NULL ||
        strstr(line, "Creating recognizer") != NULL ||
        strstr(line, "recognizer created") != NULL ||
        strstr(line, "Started") != NULL ||
        strstr(line, "Done!") != NULL ||
        strstr(line, "num threads:") != NULL ||
        strstr(line, "decoding method:") != NULL ||
        strstr(line, "Elapsed seconds:") != NULL ||
        strstr(line, "Real time factor") != NULL) {
        return TRUE;
    }

    if (strstr(line, ".wav") != NULL && (line[1] == ':' || line[0] == '/' || line[0] == '\\')) {
        return TRUE;
    }

    return FALSE;
}

static char *extract_json_text_field(const char *line) {
    const char *p = NULL;
    const char *start = NULL;
    const char *q = NULL;
    char *out = NULL;
    size_t out_len = 0;

    if (!line) {
        return NULL;
    }

    p = strstr(line, "\"text\"");
    if (!p) {
        p = strstr(line, "\"content\"");
    }
    if (!p) {
        p = strstr(line, "\"message\"");
    }
    if (!p) {
        return NULL;
    }

    p = strchr(p, ':');
    if (!p) {
        return NULL;
    }
    p++;

    while (*p == ' ' || *p == '\t') {
        p++;
    }

    if (*p != '"') {
        return NULL;
    }
    p++;
    start = p;

    q = start;
    while (*q) {
        if (*q == '"' && q > start && q[-1] != '\\') {
            break;
        }
        q++;
    }

    if (*q != '"') {
        return NULL;
    }

    out = (char *)malloc((size_t)(q - start) + 1);
    if (!out) {
        return NULL;
    }

    while (start < q) {
        if (*start == '\\' && start + 1 < q) {
            start++;
            if (*start == 'n') {
                out[out_len++] = '\n';
            } else if (*start == 'r') {
                out[out_len++] = '\r';
            } else if (*start == 't') {
                out[out_len++] = '\t';
            } else {
                out[out_len++] = *start;
            }
            start++;
            continue;
        }

        out[out_len++] = *start;
        start++;
    }

    out[out_len] = '\0';
    trim_ascii_whitespace(out);
    if (out[0] == '\0') {
        free(out);
        return NULL;
    }

    if (!seems_transcription_text(out)) {
        free(out);
        return NULL;
    }

    return out;
}

static char *extract_text_from_output(const char *output) {
    char *copy = NULL;
    char *line = NULL;
    char *save_ptr = NULL;
    char *best = NULL;

    if (!output || output[0] == '\0') {
        return NULL;
    }

    copy = _strdup(output);
    if (!copy) {
        return NULL;
    }

    line = strtok_s(copy, "\r\n", &save_ptr);
    while (line) {
        char *candidate = line;
        char *marker = NULL;
        char *json_text = NULL;

        trim_ascii_whitespace(candidate);
        if (candidate[0] == '\0') {
            line = strtok_s(NULL, "\r\n", &save_ptr);
            continue;
        }

        json_text = extract_json_text_field(candidate);
        if (json_text) {
            free(best);
            best = json_text;
            line = strtok_s(NULL, "\r\n", &save_ptr);
            continue;
        }

        marker = strstr(candidate, "Decoded text:");
        if (!marker) {
            marker = strstr(candidate, "result:");
        }
        if (!marker) {
            marker = strstr(candidate, "Result:");
        }

        if (marker) {
            marker = strchr(marker, ':');
            if (marker) {
                marker++;
                trim_ascii_whitespace(marker);
                if (marker[0] != '\0' && seems_transcription_text(marker)) {
                    free(best);
                    best = _strdup(marker);
                }
            }
        } else {
            if (!is_noisy_status_line(candidate) && seems_transcription_text(candidate)) {
                free(best);
                best = _strdup(candidate);
            }
        }

        line = strtok_s(NULL, "\r\n", &save_ptr);
    }

    free(copy);
    if (best) {
        trim_ascii_whitespace(best);
        if (best[0] == '\0' || !seems_transcription_text(best)) {
            free(best);
            best = NULL;
        }
    }

    return best;
}

static BOOL build_command_line(const wchar_t *exe_path,
                               const wchar_t *args,
                               const wchar_t *wav_path,
                               wchar_t **out_command_line) {
    size_t exe_len = 0;
    size_t args_len = 0;
    size_t wav_len = 0;
    size_t total = 0;
    wchar_t *command = NULL;

    if (!exe_path || !wav_path || !out_command_line) {
        return FALSE;
    }

    exe_len = wcslen(exe_path);
    args_len = args ? wcslen(args) : 0;
    wav_len = wcslen(wav_path);
    total = exe_len + args_len + wav_len + 16;

    command = (wchar_t *)malloc((total + 1) * sizeof(wchar_t));
    if (!command) {
        return FALSE;
    }

    if (args_len > 0) {
        swprintf(command, total + 1, L"\"%ls\" %ls \"%ls\"", exe_path, args, wav_path);
    } else {
        swprintf(command, total + 1, L"\"%ls\" \"%ls\"", exe_path, wav_path);
    }

    *out_command_line = command;
    return TRUE;
}

AsrBackendKind asr_parse_backend_name(const wchar_t *name) {
    if (!name || name[0] == L'\0') {
        return ASR_BACKEND_GROQ;
    }

    if (_wcsicmp(name, L"sherpa") == 0 || _wcsicmp(name, L"local") == 0) {
        return ASR_BACKEND_SHERPA;
    }
    if (_wcsicmp(name, L"gladia") == 0) {
        return ASR_BACKEND_GLADIA;
    }
    if (_wcsicmp(name, L"gemini") == 0) {
        return ASR_BACKEND_GEMINI;
    }

    return ASR_BACKEND_GROQ;
}

const wchar_t *asr_backend_name(AsrBackendKind backend) {
    if (backend == ASR_BACKEND_SHERPA) {
        return L"sherpa";
    }
    if (backend == ASR_BACKEND_GLADIA) {
        return L"gladia";
    }
    if (backend == ASR_BACKEND_GEMINI) {
        return L"gemini";
    }

    return L"groq";
}

BOOL sherpa_transcribe_wav_cli(const wchar_t *wav_path,
                               const wchar_t *sherpa_exe,
                               const wchar_t *sherpa_args,
                               char **out_utf8_text,
                               char **out_error_utf8) {
    const wchar_t *exe_path = sherpa_exe;
    wchar_t *command_line = NULL;
    wchar_t working_dir[MAX_PATH];

    SECURITY_ATTRIBUTES sec_attr;
    HANDLE read_pipe = NULL;
    HANDLE write_pipe = NULL;
    STARTUPINFOW startup;
    PROCESS_INFORMATION process;

    BOOL ok = FALSE;
    char *output = NULL;
    size_t output_len = 0;
    DWORD exit_code = 0;

    if (!wav_path || !out_utf8_text) {
        set_error_text(out_error_utf8, "Invalid arguments for sherpa backend");
        return FALSE;
    }

    *out_utf8_text = NULL;
    if (out_error_utf8) {
        free(*out_error_utf8);
        *out_error_utf8 = NULL;
    }

    if (is_empty_wide(exe_path)) {
        exe_path = L"sherpa-onnx-offline.exe";
    }

    if (!build_command_line(exe_path, sherpa_args, wav_path, &command_line)) {
        set_error_text(out_error_utf8, "Failed to build sherpa command line");
        return FALSE;
    }

    // 调试：记录启动参数，排查乱码和路径长度问题
    // 注意：这里由于是 asr_backend，不直接引用 app 指针，建议此处根据需要打印或返回
    // 我们暂时在 CreateProcessW 处打印

    ZeroMemory(&sec_attr, sizeof(sec_attr));
    sec_attr.nLength = sizeof(sec_attr);
    sec_attr.bInheritHandle = TRUE;

    if (!CreatePipe(&read_pipe, &write_pipe, &sec_attr, 0)) {
        set_error_text(out_error_utf8, "Failed to create pipe for sherpa output");
        goto cleanup;
    }

    if (!SetHandleInformation(read_pipe, HANDLE_FLAG_INHERIT, 0)) {
        set_error_text(out_error_utf8, "Failed to set pipe handle information");
        goto cleanup;
    }

    ZeroMemory(&startup, sizeof(startup));
    startup.cb = sizeof(startup);
    startup.dwFlags = STARTF_USESTDHANDLES;
    startup.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
    startup.hStdOutput = write_pipe;
    startup.hStdError = write_pipe;

    ZeroMemory(&process, sizeof(process));
    ZeroMemory(working_dir, sizeof(working_dir));
    extract_directory(exe_path, working_dir, _countof(working_dir));

    if (!CreateProcessW(NULL,
                        command_line,
                        NULL,
                        NULL,
                        TRUE,
                        CREATE_NO_WINDOW,
                        NULL,
                        working_dir[0] ? working_dir : NULL,
                        &startup,
                        &process)) {
        char message[512];
        snprintf(message, sizeof(message),
                 "Failed to start sherpa executable. Ensure path is valid: %ls", exe_path);
        set_error_text(out_error_utf8, message);
        goto cleanup;
    }

    CloseHandle(write_pipe);
    write_pipe = NULL;

    for (;;) {
        char buffer[2048];
        DWORD bytes_read = 0;

        if (!ReadFile(read_pipe, buffer, sizeof(buffer), &bytes_read, NULL)) {
            if (GetLastError() == ERROR_BROKEN_PIPE) {
                break;
            }
            set_error_text(out_error_utf8, "Failed while reading sherpa output");
            goto cleanup;
        }

        if (bytes_read == 0) {
            break;
        }

        if (!append_chunk(&output, &output_len, buffer, bytes_read)) {
            set_error_text(out_error_utf8, "Sherpa output is too large");
            goto cleanup;
        }
    }

    WaitForSingleObject(process.hProcess, INFINITE);
    GetExitCodeProcess(process.hProcess, &exit_code);

    if (!output) {
        output = _strdup("");
        if (!output) {
            set_error_text(out_error_utf8, "Out of memory while collecting sherpa output");
            goto cleanup;
        }
    }

    if (exit_code != 0) {
        char message[256];
        snprintf(message, sizeof(message), "Sherpa process exited with code %lu", (unsigned long)exit_code);
        set_error_text(out_error_utf8, message);
        if (out_error_utf8 && output[0] != '\0') {
            char *detailed = NULL;
            size_t prefix_len = strlen(*out_error_utf8);
            size_t output_keep = strlen(output);
            if (output_keep > 1000) {
                output_keep = 1000;
            }

            detailed = (char *)malloc(prefix_len + output_keep + 4);
            if (detailed) {
                memcpy(detailed, *out_error_utf8, prefix_len);
                detailed[prefix_len] = ':';
                detailed[prefix_len + 1] = ' ';
                memcpy(detailed + prefix_len + 2, output, output_keep);
                detailed[prefix_len + 2 + output_keep] = '\0';
                free(*out_error_utf8);
                *out_error_utf8 = detailed;
            }
        }
        goto cleanup;
    }

    *out_utf8_text = extract_text_from_output(output);
    if (!*out_utf8_text) {
        set_error_text(out_error_utf8, "Sherpa finished but did not return transcription text");
        goto cleanup;
    }

    ok = TRUE;

cleanup:
    if (process.hThread) {
        CloseHandle(process.hThread);
    }
    if (process.hProcess) {
        CloseHandle(process.hProcess);
    }

    if (read_pipe) {
        CloseHandle(read_pipe);
    }
    if (write_pipe) {
        CloseHandle(write_pipe);
    }

    free(command_line);
    free(output);
    return ok;
}

BOOL sherpa_transcribe_wav_websocket(const wchar_t *wav_path,
                                     char **out_utf8_text,
                                     char **out_error_utf8) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    LARGE_INTEGER file_size = {0};
    unsigned char *wav_data = NULL;
    DWORD bytes_read = 0;
    BOOL ok = FALSE;

    HINTERNET session = NULL;
    HINTERNET connect = NULL;
    HINTERNET request = NULL;
    HINTERNET web_socket = NULL;

    if (!wav_path || !out_utf8_text) {
        set_error_text(out_error_utf8, "Invalid arguments for websocket transcriber");
        return FALSE;
    }

    *out_utf8_text = NULL;
    if (out_error_utf8) {
        free(*out_error_utf8);
        *out_error_utf8 = NULL;
    }

    // 1. Read WAV File
    file_handle = CreateFileW(wav_path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file_handle == INVALID_HANDLE_VALUE) {
        set_error_text(out_error_utf8, "Failed to open WAV file");
        return FALSE;
    }

    if (!GetFileSizeEx(file_handle, &file_size) || file_size.QuadPart <= 44 || file_size.QuadPart > 50 * 1024 * 1024) {
        CloseHandle(file_handle);
        set_error_text(out_error_utf8, "WAV file is invalid or too large");
        return FALSE;
    }

    wav_data = (unsigned char *)malloc((size_t)file_size.QuadPart);
    if (!wav_data) {
        CloseHandle(file_handle);
        set_error_text(out_error_utf8, "Out of memory reading WAV file");
        return FALSE;
    }

    if (!ReadFile(file_handle, wav_data, (DWORD)file_size.QuadPart, &bytes_read, NULL) || bytes_read != (DWORD)file_size.QuadPart) {
        free(wav_data);
        CloseHandle(file_handle);
        set_error_text(out_error_utf8, "Failed to read WAV bytes");
        return FALSE;
    }
    CloseHandle(file_handle);

    // Parse WAV Header (WAV fmt mono 16kHz 16-bit PCM has samples starting at byte 44)
    int32_t sample_rate = *(int32_t *)(wav_data + 24);
    int16_t bits_per_sample = *(int16_t *)(wav_data + 34);

    if (bits_per_sample != 16) {
        free(wav_data);
        set_error_text(out_error_utf8, "Only 16-bit PCM WAV is supported by this client");
        return FALSE;
    }

    DWORD pcm_bytes = bytes_read - 44;
    DWORD num_samples = pcm_bytes / sizeof(int16_t);
    int16_t *pcm_samples = (int16_t *)(wav_data + 44);

    // Convert to Float32 samples scaled to [-1.0, 1.0]
    float *float_samples = (float *)malloc(num_samples * sizeof(float));
    if (!float_samples) {
        free(wav_data);
        set_error_text(out_error_utf8, "Out of memory allocating float buffer");
        return FALSE;
    }

    for (DWORD i = 0; i < num_samples; ++i) {
        float_samples[i] = (float)pcm_samples[i] / 32768.0f;
    }
    free(wav_data);

    // 2. Prepare WebSocket Binary Package
    // Package format: [sample_rate (int32)][num_audio_bytes (int32)][samples (float32...)]
    DWORD num_audio_bytes = num_samples * sizeof(float);
    DWORD payload_size = 4 + 4 + num_audio_bytes;
    unsigned char *payload = (unsigned char *)malloc(payload_size);
    if (!payload) {
        free(float_samples);
        set_error_text(out_error_utf8, "Out of memory allocating payload");
        return FALSE;
    }

    *(int32_t *)(payload + 0) = sample_rate;
    *(int32_t *)(payload + 4) = (int32_t)num_audio_bytes;
    memcpy(payload + 8, float_samples, num_audio_bytes);
    free(float_samples);

    // 3. Setup WinHTTP & WebSocket upgrade handshake with auto-retry for slow booting machines
    int retries = 10;
    BOOL connected = FALSE;
    for (int i = 0; i < retries; ++i) {
        session = WinHttpOpen(L"VoiceIME/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                              WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (!session) {
            Sleep(500);
            continue;
        }

        DWORD timeout = 1000;
        WinHttpSetOption(session, WINHTTP_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));

        connect = WinHttpConnect(session, L"127.0.0.1", 6006, 0);
        if (!connect) {
            WinHttpCloseHandle(session);
            session = NULL;
            Sleep(500);
            continue;
        }

        request = WinHttpOpenRequest(connect, L"GET", L"/", NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
        if (!request) {
            WinHttpCloseHandle(connect);
            connect = NULL;
            WinHttpCloseHandle(session);
            session = NULL;
            Sleep(500);
            continue;
        }

        if (!WinHttpSetOption(request, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, NULL, 0)) {
            WinHttpCloseHandle(request);
            request = NULL;
            WinHttpCloseHandle(connect);
            connect = NULL;
            WinHttpCloseHandle(session);
            session = NULL;
            Sleep(500);
            continue;
        }

        if (WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
            connected = TRUE;
            break;
        }

        WinHttpCloseHandle(request);
        request = NULL;
        WinHttpCloseHandle(connect);
        connect = NULL;
        WinHttpCloseHandle(session);
        session = NULL;
        
        Sleep(500); // Wait 500ms before retrying
    }

    if (!connected) {
        free(payload);
        set_error_text(out_error_utf8, "Local WebSocket server is not running on port 6006.");
        return FALSE;
    }

    if (!WinHttpReceiveResponse(request, NULL)) {
        free(payload);
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        set_error_text(out_error_utf8, "Failed to receive handshake response");
        return FALSE;
    }

    web_socket = WinHttpWebSocketCompleteUpgrade(request, 0);
    if (!web_socket) {
        free(payload);
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        set_error_text(out_error_utf8, "WebSocket upgrade could not be completed");
        return FALSE;
    }
    WinHttpCloseHandle(request);
    request = NULL;

    // 4. Send Binary Audio Payload
    DWORD ws_err = WinHttpWebSocketSend(web_socket, WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, payload, payload_size);
    free(payload);
    if (ws_err != ERROR_SUCCESS) {
        char err_msg[128];
        snprintf(err_msg, sizeof(err_msg), "WinHttpWebSocketSend failed with error %lu", (unsigned long)ws_err);
        set_error_text(out_error_utf8, err_msg);
        goto cleanup;
    }

    // 5. Receive Response JSON
    char *response_buffer = NULL;
    size_t response_capacity = 4096;
    size_t response_len = 0;
    response_buffer = (char *)malloc(response_capacity);
    if (!response_buffer) {
        set_error_text(out_error_utf8, "Out of memory allocating response buffer");
        goto cleanup;
    }
    response_buffer[0] = '\0';

    for (;;) {
        BYTE chunk[2048];
        DWORD bytes_transferred = 0;
        WINHTTP_WEB_SOCKET_BUFFER_TYPE buffer_type;

        ws_err = WinHttpWebSocketReceive(web_socket, chunk, sizeof(chunk), &bytes_transferred, &buffer_type);
        if (ws_err != ERROR_SUCCESS) {
            char err_msg[128];
            snprintf(err_msg, sizeof(err_msg), "WinHttpWebSocketReceive failed with error %lu", (unsigned long)ws_err);
            set_error_text(out_error_utf8, err_msg);
            free(response_buffer);
            goto cleanup;
        }

        if (bytes_transferred == 0) {
            break;
        }

        if (response_len + bytes_transferred >= response_capacity) {
            response_capacity = response_len + bytes_transferred + 4096;
            char *new_buf = (char *)realloc(response_buffer, response_capacity);
            if (!new_buf) {
                set_error_text(out_error_utf8, "Out of memory growing response buffer");
                free(response_buffer);
                goto cleanup;
            }
            response_buffer = new_buf;
        }

        memcpy(response_buffer + response_len, chunk, bytes_transferred);
        response_len += bytes_transferred;
        response_buffer[response_len] = '\0';

        if (buffer_type == WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE ||
            buffer_type == WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE) {
            break;
        }
    }

    // 6. Send "Done" text and close WebSocket connection
    WinHttpWebSocketSend(web_socket, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, (void *)"Done", 4);
    WinHttpWebSocketClose(web_socket, WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, NULL, 0);

    // 7. Parse the text field from JSON response
    if (response_len > 0) {
        *out_utf8_text = extract_json_text_field(response_buffer);
        if (!*out_utf8_text) {
            *out_utf8_text = response_buffer;
            response_buffer = NULL;
        }
        ok = TRUE;
    } else {
        set_error_text(out_error_utf8, "WebSocket server returned empty result");
    }
    free(response_buffer);

    // Trim trailing quotes or newlines if any
    if (ok && *out_utf8_text) {
        trim_ascii_whitespace(*out_utf8_text);
    }

cleanup:
    if (web_socket) {
        WinHttpCloseHandle(web_socket);
    }
    if (connect) {
        WinHttpCloseHandle(connect);
    }
    if (session) {
        WinHttpCloseHandle(session);
    }
    return ok;
}

static void get_exe_sibling_path(const wchar_t *filename, wchar_t *out_path, size_t out_cap) {
    GetModuleFileNameW(NULL, out_path, (DWORD)out_cap);
    wchar_t *last_slash = wcsrchr(out_path, L'\\');
    if (last_slash) {
        *(last_slash + 1) = L'\0';
    }
    
    // Create prompts/ subdirectory
    wchar_t prompts_dir[MAX_PATH];
    wcscpy_s(prompts_dir, _countof(prompts_dir), out_path);
    wcscat_s(prompts_dir, _countof(prompts_dir), L"prompts");
    CreateDirectoryW(prompts_dir, NULL);
    
    wcscat_s(out_path, out_cap, L"prompts\\");
    wcscat_s(out_path, out_cap, filename);
}

static char *load_system_prompt(const wchar_t *filename, const char *default_prompt_utf8) {
    wchar_t path[MAX_PATH];
    get_exe_sibling_path(filename, path, _countof(path));

    HANDLE file = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        file = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (file != INVALID_HANDLE_VALUE) {
            DWORD written = 0;
            WriteFile(file, default_prompt_utf8, (DWORD)strlen(default_prompt_utf8), &written, NULL);
            CloseHandle(file);
        }
        return _strdup(default_prompt_utf8);
    }

    DWORD file_size = GetFileSize(file, NULL);
    if (file_size == INVALID_FILE_SIZE || file_size == 0) {
        CloseHandle(file);
        return _strdup(default_prompt_utf8);
    }

    char *buffer = (char *)malloc(file_size + 1);
    if (!buffer) {
        CloseHandle(file);
        return _strdup(default_prompt_utf8);
    }

    DWORD read = 0;
    if (ReadFile(file, buffer, file_size, &read, NULL)) {
        buffer[read] = '\0';
    } else {
        free(buffer);
        buffer = _strdup(default_prompt_utf8);
    }
    CloseHandle(file);

    char *result = buffer;
    if ((unsigned char)buffer[0] == 0xEF && (unsigned char)buffer[1] == 0xBB && (unsigned char)buffer[2] == 0xBF) {
        result = _strdup(buffer + 3);
        free(buffer);
    }
    return result;
}

static void append_fewshot_messages_json(char *dest, size_t dest_cap, const char *fewshot_text) {
    if (!fewshot_text || fewshot_text[0] == '\0') {
        return;
    }
    
    char *copy = _strdup(fewshot_text);
    if (!copy) return;
    
    char *line = NULL;
    char *save_ptr = NULL;
    line = strtok_s(copy, "\r\n", &save_ptr);
    
    while (line) {
        while (*line == ' ' || *line == '\t') {
            line++;
        }
        
        const char *role = NULL;
        const char *content_start = NULL;
        
        if (_strnicmp(line, "user:", 5) == 0) {
            role = "user";
            content_start = line + 5;
        } else if (strncmp(line, "user\xEF\xBC\x9A", 7) == 0) {
            role = "user";
            content_start = line + 7;
        } else if (_strnicmp(line, "assistant:", 10) == 0) {
            role = "assistant";
            content_start = line + 10;
        } else if (strncmp(line, "assistant\xEF\xBC\x9A", 12) == 0) {
            role = "assistant";
            content_start = line + 12;
        }
        
        if (role && content_start) {
            while (*content_start == ' ' || *content_start == '\t') {
                content_start++;
            }
            
            char *content_copy = _strdup(content_start);
            if (content_copy) {
                size_t len = strlen(content_copy);
                while (len > 0 && (content_copy[len - 1] == ' ' || content_copy[len - 1] == '\t')) {
                    content_copy[len - 1] = '\0';
                    len--;
                }
                
                char *escaped = json_escape(content_copy);
                free(content_copy);
                
                if (escaped) {
                    char temp[4096];
                    snprintf(temp, sizeof(temp), "{\"role\":\"%s\",\"content\":\"%s\"},", role, escaped);
                    free(escaped);
                    
                    if (strlen(dest) + strlen(temp) < dest_cap) {
                        strcat_s(dest, dest_cap, temp);
                    }
                }
            }
        }
        
        line = strtok_s(NULL, "\r\n", &save_ptr);
    }
    
    free(copy);
}

static void autocomplete_api_url(const wchar_t *input, wchar_t *output, size_t output_cap) {
    wcsncpy_s(output, output_cap, input, _TRUNCATE);
    
    if (wcsstr(output, L"/api/v1/chat") != NULL) {
        return;
    }
    
    size_t len = wcslen(output);
    if (len >= 17 && _wcsnicmp(output + len - 17, L"/chat/completions", 17) == 0) {
        return;
    }
    
    if (wcsstr(output, L"/v1") == NULL) {
        if (len > 0 && output[len - 1] == L'/') {
            wcscat_s(output, output_cap, L"v1/chat/completions");
        } else {
            wcscat_s(output, output_cap, L"/v1/chat/completions");
        }
    } else {
        if (len > 0 && output[len - 1] == L'/') {
            wcscat_s(output, output_cap, L"chat/completions");
        } else {
            wcscat_s(output, output_cap, L"/chat/completions");
        }
    }
}

static BOOL parse_url_components(const wchar_t *url_str,
                                 wchar_t *out_host, size_t host_cap,
                                 INTERNET_PORT *out_port,
                                 wchar_t *out_path, size_t path_cap,
                                 BOOL *out_is_secure) {
    URL_COMPONENTS urlComp;
    ZeroMemory(&urlComp, sizeof(urlComp));
    urlComp.dwStructSize = sizeof(urlComp);
    
    urlComp.lpszHostName = out_host;
    urlComp.dwHostNameLength = (DWORD)host_cap;
    
    urlComp.lpszUrlPath = out_path;
    urlComp.dwUrlPathLength = (DWORD)path_cap;
    
    if (WinHttpCrackUrl(url_str, (DWORD)wcslen(url_str), 0, &urlComp)) {
        *out_port = urlComp.nPort;
        *out_is_secure = (urlComp.nScheme == INTERNET_SCHEME_HTTPS);
        return TRUE;
    }
    return FALSE;
}

static char *json_escape(const char *input) {
    if (!input) return NULL;
    size_t len = strlen(input);
    size_t out_capacity = len * 2 + 1;
    char *output = (char *)malloc(out_capacity);
    if (!output) return NULL;

    size_t j = 0;
    for (size_t i = 0; i < len; ++i) {
        char c = input[i];
        if (c == '\\' || c == '"') {
            output[j++] = '\\';
            output[j++] = c;
        } else if (c == '\n') {
            output[j++] = '\\';
            output[j++] = 'n';
        } else if (c == '\r') {
            output[j++] = '\\';
            output[j++] = 'r';
        } else if (c == '\t') {
            output[j++] = '\\';
            output[j++] = 't';
        } else {
            output[j++] = c;
        }
    }
    output[j] = '\0';
    return output;
}

static char *build_openai_chat_payload(const char *input_text,
                                       const char *model_name,
                                       const char *target_lang,
                                       const char *style,
                                       const char *custom_prompt,
                                       const char *replace_rules,
                                       const char *ai_vocab) {
    char *style_loaded = NULL;
    if (strcmp(style, "商務正式") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_business.txt", "请将文本改写得更加商务、正式、专业，适合职场与商务邮件沟通。");
    } else if (strcmp(style, "日常口語") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_casual.txt", "请保持日常口语风格，使语句流畅自然，不要过于死板。");
    } else if (strcmp(style, "簡潔扼要") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_concise.txt", "请尽可能简洁明了，去掉赘字，保留核心重点。");
    } else {
        style_loaded = load_system_prompt(L"prompt_style_default.txt", "请优化语句，使其流畅、重点清晰。");
    }

    char style_instruction[1024] = "";
    if (style_loaded) {
        strncpy_s(style_instruction, sizeof(style_instruction), style_loaded, _TRUNCATE);
        free(style_loaded);
    }

    char lang_instruction[256] = "";
    if (strcmp(target_lang, "不翻譯") != 0 && strlen(target_lang) > 0) {
        sprintf_s(lang_instruction, sizeof(lang_instruction), "并将最终结果翻译为「%s」（如果是中文请用对应的繁简体）。", target_lang);
    }

    char dict_instruction[2048] = "";
    if (replace_rules && strlen(replace_rules) > 0) {
        sprintf_s(dict_instruction, sizeof(dict_instruction), 
                  "\n纠错词典（将谐音错字更正为）：\n%s\n", 
                  replace_rules);
    }

    char vocab_instruction[2048] = "";
    if (ai_vocab && strlen(ai_vocab) > 0) {
        sprintf_s(vocab_instruction, sizeof(vocab_instruction), 
                  "\n【专有名词与自订词库（若原始文字中听起来相近，请优先使用且纠正为此处的写法）】:\n%s\n", 
                  ai_vocab);
    }

    char custom_instruction[1024] = "";
    if (custom_prompt && strlen(custom_prompt) > 0 && strcmp(style, "自訂 Prompt") == 0) {
        sprintf_s(custom_instruction, sizeof(custom_instruction), "\n【自订额外要求】:\n%s\n", custom_prompt);
    }

    const char *default_openai_prompt = 
        "你是一个极其简练的语音输入法后台文本优化助手。你的任务是对用户的语音识别文本内容进行正确的加标点、去填充词（如：嗯、呃、那个）和重复句子修正。\n"
        "注意遵守规则：\n"
        "1. 你必须且只能直接输出优化后的文本，绝对不能包含任何问候、解释、分析、对话或回答文本中的内容。\n"
        "2. 英文缩写拼写修正：必须把被拆开的英文缩写合并并转为大写（如：把 'a i' 合并为 'AI'，把 'i p' 合并为 'IP'，把 'a p i' 合并为 'API'），绝对不能把 'a i' 中的字母 'a' 误判为语气词 '啊/a' 而删除！";

    char *base_prompt = load_system_prompt(L"system_prompt_openai.txt", default_openai_prompt);
    char *system_prompt = NULL;
    if (base_prompt) {
        size_t system_prompt_cap = strlen(base_prompt) + strlen(style_instruction) + strlen(lang_instruction) + 
                                   strlen(dict_instruction) + strlen(vocab_instruction) + strlen(custom_instruction) + 256;
        system_prompt = (char *)malloc(system_prompt_cap);
        if (system_prompt) {
            sprintf_s(system_prompt, system_prompt_cap,
                      "%s\n3. %s%s%s%s%s",
                      base_prompt,
                      style_instruction,
                      lang_instruction,
                      dict_instruction,
                      vocab_instruction,
                      custom_instruction);
        }
        free(base_prompt);
    }

    if (!system_prompt) {
        return NULL;
    }

    char *escaped_system = json_escape(system_prompt);
    free(system_prompt);
    char *escaped_input = json_escape(input_text);
    char *escaped_model = json_escape(model_name);

    if (!escaped_system || !escaped_input || !escaped_model) {
        free(escaped_system);
        free(escaped_input);
        free(escaped_model);
        return NULL;
    }

    size_t msgs_cap = strlen(escaped_system) + strlen(escaped_input) + 32768;
    char *msgs_json = (char *)malloc(msgs_cap);
    if (!msgs_json) {
        free(escaped_system);
        free(escaped_input);
        free(escaped_model);
        return NULL;
    }
    msgs_json[0] = '\0';

    char sys_msg_header[512] = "{\"role\":\"system\",\"content\":\"";
    char sys_msg_footer[32] = "\"},";
    if (strlen(msgs_json) + strlen(sys_msg_header) + strlen(escaped_system) + strlen(sys_msg_footer) < msgs_cap) {
        strcat_s(msgs_json, msgs_cap, sys_msg_header);
        strcat_s(msgs_json, msgs_cap, escaped_system);
        strcat_s(msgs_json, msgs_cap, sys_msg_footer);
    }

    const char *default_openai_fewshot = 
        "user: 那个今天天气呃挺好的吧\n"
        "assistant: 今天天气挺好的吧。\n"
        "user: 这个是结合 a i 自动纠错比那个更智能\n"
        "assistant: 这个是结合 AI 自动纠错，比那个更智能。\n"
        "user: 你是谁呀\n"
        "assistant: 你是谁呀。\n"
        "user: 不换嘛那你得当时说一下看看才知道\n"
        "assistant: 不换嘛，那你得当时说一下，看看才知道。\n";

    char *fewshot_prompt = load_system_prompt(L"prompt_openai_fewshot.txt", default_openai_fewshot);
    if (fewshot_prompt) {
        append_fewshot_messages_json(msgs_json, msgs_cap, fewshot_prompt);
        free(fewshot_prompt);
    }

    char user_msg[4096];
    snprintf(user_msg, sizeof(user_msg), "{\"role\":\"user\",\"content\":\"待优化的原始语音文本：\\\"%s\\\"\"}", escaped_input);
    if (strlen(msgs_json) + strlen(user_msg) < msgs_cap) {
        strcat_s(msgs_json, msgs_cap, user_msg);
    }

    size_t payload_len = strlen(escaped_model) + strlen(msgs_json) + 4096;
    char *payload = (char *)malloc(payload_len);
    if (!payload) {
        free(msgs_json);
        free(escaped_system);
        free(escaped_input);
        free(escaped_model);
        return NULL;
    }

    sprintf_s(payload, payload_len,
              "{"
              "\"model\":\"%s\","
              "\"messages\":[%s],"
              "\"temperature\":0.2"
              "}",
              escaped_model,
              msgs_json);

    free(msgs_json);
    free(escaped_system);
    free(escaped_input);
    free(escaped_model);
    return payload;
}

static char *build_polish_prompt(const char *input_text,
                                 const char *target_lang,
                                 const char *style,
                                 const char *custom_prompt,
                                 const char *replace_rules,
                                 const char *ai_vocab) {
    char *style_loaded = NULL;
    if (strcmp(style, "商務正式") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_business.txt", "请将文本改写得更加商务、正式、专业，适合职场与商务邮件沟通。");
    } else if (strcmp(style, "日常口語") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_casual.txt", "请保持日常口语风格，使语句流畅自然，不要过于死板。");
    } else if (strcmp(style, "簡潔扼要") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_concise.txt", "请尽可能简洁明了，去掉赘字，保留核心重点。");
    } else {
        style_loaded = load_system_prompt(L"prompt_style_default.txt", "请优化语句，使其流畅、重点清晰。");
    }

    char style_instruction[1024] = "";
    if (style_loaded) {
        strncpy_s(style_instruction, sizeof(style_instruction), style_loaded, _TRUNCATE);
        free(style_loaded);
    }

    char lang_instruction[256] = "";
    if (strcmp(target_lang, "不翻譯") != 0 && strlen(target_lang) > 0) {
        sprintf_s(lang_instruction, sizeof(lang_instruction), "并将最终结果翻译为「%s」（如果是中文请用对应的繁简体）。", target_lang);
    }

    char dict_instruction[2048] = "";
    if (replace_rules && strlen(replace_rules) > 0) {
        sprintf_s(dict_instruction, sizeof(dict_instruction), 
                  "\n【专有名词与字词拼写参考（若语音中听起来相近，请优先使用以下写法）】:\n%s\n", 
                  replace_rules);
    }

    char vocab_instruction[2048] = "";
    if (ai_vocab && strlen(ai_vocab) > 0) {
        sprintf_s(vocab_instruction, sizeof(vocab_instruction), 
                  "\n【专有名词与自订词库（若原始文字中听起来相近，请优先使用且纠正为此处的写法）】:\n%s\n", 
                  ai_vocab);
    }

    char custom_instruction[1024] = "";
    if (custom_prompt && strlen(custom_prompt) > 0 && strcmp(style, "自訂 Prompt") == 0) {
        sprintf_s(custom_instruction, sizeof(custom_instruction), "\n【自订额外要求】:\n%s\n", custom_prompt);
    }

    const char *default_gemini_prompt = 
        "你是一个专业的语音输入优化助理。请对以下【原始转写文字】进行智能优化、重写与排版。\n"
        "【核心任务】:\n"
        "1. 移除无意义的口头语、填充词（如“嗯”、“呃”、“啊”、“然后”、“那个”、“你知道的”等）。\n"
        "2. 移除不必要的重复词汇与结巴。\n"
        "3. 自动识别并处理自我修正（例如“不对，应该上午十点” -> 直接改为“上午十点”）。\n"
        "4. 自动将口述的清单、步骤或要点整理成干净、结构化的分行列表或标点符号排版。\n"
        "【输出规则】:\n"
        "- 请“仅”输出优化后的最终文本，不要包含任何解释、Markdown代码块标记、导言或括号说明。\n"
        "- 如果原始文字自动识别为无意义的噪音、空字串或无法理清的语音，请直接输出空字串。";

    char *base_prompt = load_system_prompt(L"system_prompt_gemini.txt", default_gemini_prompt);
    if (!base_prompt) {
        return NULL;
    }

    size_t total_len = strlen(base_prompt) + strlen(input_text) + strlen(style_instruction) + strlen(lang_instruction) +
                       strlen(dict_instruction) + strlen(vocab_instruction) + strlen(custom_instruction) + 2048;
    char *prompt = (char *)malloc(total_len);
    if (!prompt) {
        free(base_prompt);
        return NULL;
    }

    sprintf_s(prompt, total_len,
              "%s\n\n【额外指示】:\n5. %s%s%s%s%s\n\n【原始转写文字】:\n\"%s\"",
              base_prompt,
              style_instruction,
              lang_instruction,
              dict_instruction,
              vocab_instruction,
              custom_instruction,
              input_text);
    free(base_prompt);

    return prompt;
}

static char *build_transcribe_prompt(const char *target_lang,
                                     const char *style,
                                     const char *custom_prompt,
                                     const char *replace_rules,
                                     const char *ai_vocab) {
    char *style_loaded = NULL;
    if (strcmp(style, "商務正式") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_business.txt", "请将文本改写得更加商务、正式、专业，适合职场与商务邮件沟通。");
    } else if (strcmp(style, "日常口語") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_casual.txt", "请保持日常口语风格，使语句流畅自然，不要过于死板。");
    } else if (strcmp(style, "簡潔扼要") == 0) {
        style_loaded = load_system_prompt(L"prompt_style_concise.txt", "请尽可能简洁明了，去掉赘字，保留核心重点。");
    } else {
        style_loaded = load_system_prompt(L"prompt_style_default.txt", "请优化语句，使其流畅、重点清晰。");
    }

    char style_instruction[1024] = "";
    if (style_loaded) {
        strncpy_s(style_instruction, sizeof(style_instruction), style_loaded, _TRUNCATE);
        free(style_loaded);
    }

    char lang_instruction[256] = "";
    if (strcmp(target_lang, "不翻譯") != 0 && strlen(target_lang) > 0) {
        sprintf_s(lang_instruction, sizeof(lang_instruction), "并将结果翻译为「%s」。", target_lang);
    }

    char dict_instruction[2048] = "";
    if (replace_rules && strlen(replace_rules) > 0) {
        sprintf_s(dict_instruction, sizeof(dict_instruction), 
                  "\n【专有名词与字词拼写参考（若语音中发音相近，请优先使用以下写法）】:\n%s\n", 
                  replace_rules);
    }

    char vocab_instruction[2048] = "";
    if (ai_vocab && strlen(ai_vocab) > 0) {
        sprintf_s(vocab_instruction, sizeof(vocab_instruction), 
                  "\n【专有名词与自订词库（若原始文字中听起来相近，请优先使用且纠正为此处的写法）】:\n%s\n", 
                  ai_vocab);
    }

    char custom_instruction[1024] = "";
    if (custom_prompt && strlen(custom_prompt) > 0 && strcmp(style, "自訂 Prompt") == 0) {
        sprintf_s(custom_instruction, sizeof(custom_instruction), "\n【自订额外要求】:\n%s\n", custom_prompt);
    }

    const char *default_gemini_transcribe_prompt = 
        "请将附带的语音直接转录为文字，并同时进行智能优化、重写与排版。\n"
        "【核心任务】:\n"
        "1. 自动去除口头赘词与无意义填充词（如“嗯”、“呃”、“啊”、“然后”、“那个”、“你知道的”等）。\n"
        "2. 消除语音中的结巴与重复词汇。\n"
        "3. 自动识别并处理自我修正（只保留修正后的最终意图，不要保留口误）。\n"
        "4. 自动将口述清单、步骤或重点整理为干净分行列表或合适的标点符号排版。";

    char *base_prompt = load_system_prompt(L"system_prompt_gemini_transcribe.txt", default_gemini_transcribe_prompt);
    if (!base_prompt) return NULL;

    size_t total_len = strlen(base_prompt) + strlen(style_instruction) + strlen(lang_instruction) +
                       strlen(dict_instruction) + strlen(vocab_instruction) + strlen(custom_instruction) + 2048;
    char *prompt = (char *)malloc(total_len);
    if (!prompt) {
        free(base_prompt);
        return NULL;
    }

    sprintf_s(prompt, total_len,
              "%s\n\n【额外指示】:\n5. %s%s%s%s%s\n\n"
              "【输出规则】:\n"
              "- 请“仅”输出优化后的转录文字，不要包含任何解释、Markdown代码块标记、导言说明。\n"
              "- 若音频中只有噪音，请输出空字串。",
              base_prompt,
              style_instruction,
              lang_instruction,
              dict_instruction,
              vocab_instruction,
              custom_instruction);

    free(base_prompt);
    return prompt;
}

static BOOL gemini_send_request(const char *api_key_utf8,
                                const char *model_utf8,
                                const char *custom_url_utf8,
                                const char *json_payload,
                                char **out_response_utf8,
                                char **out_error_utf8) {
    HINTERNET session = NULL;
    HINTERNET connect = NULL;
    HINTERNET request = NULL;
    BOOL success = FALSE;
    
    wchar_t *host_wide = NULL;
    wchar_t *path_wide = NULL;
    INTERNET_PORT port = INTERNET_DEFAULT_HTTPS_PORT;
    BOOL is_secure = TRUE;
    
    if (out_response_utf8) *out_response_utf8 = NULL;
    if (out_error_utf8) *out_error_utf8 = NULL;

    session = WinHttpOpen(L"VoiceIME/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) {
        set_error_text(out_error_utf8, "Failed to create HTTP session");
        return FALSE;
    }

    DWORD timeout = 25000;
    WinHttpSetOption(session, WINHTTP_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));
    WinHttpSetOption(session, WINHTTP_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
    WinHttpSetOption(session, WINHTTP_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));

    if (custom_url_utf8 && custom_url_utf8[0] != '\0') {
        int url_len = MultiByteToWideChar(CP_UTF8, 0, custom_url_utf8, -1, NULL, 0);
        wchar_t *custom_url_wide = (wchar_t *)malloc(url_len * sizeof(wchar_t));
        if (!custom_url_wide) {
            set_error_text(out_error_utf8, "Memory allocation failed for URL");
            goto cleanup;
        }
        MultiByteToWideChar(CP_UTF8, 0, custom_url_utf8, -1, custom_url_wide, url_len);
        
        wchar_t autocompleted_url[2048];
        autocomplete_api_url(custom_url_wide, autocompleted_url, _countof(autocompleted_url));
        free(custom_url_wide);
        
        host_wide = (wchar_t *)malloc(2048 * sizeof(wchar_t));
        path_wide = (wchar_t *)malloc(4096 * sizeof(wchar_t));
        if (!host_wide || !path_wide) {
            set_error_text(out_error_utf8, "Memory allocation failed for URL components");
            goto cleanup;
        }
        
        if (!parse_url_components(autocompleted_url, host_wide, 2048, &port, path_wide, 4096, &is_secure)) {
            set_error_text(out_error_utf8, "Failed to parse custom API URL");
            goto cleanup;
        }
    } else {
        host_wide = _wcsdup(L"generativelanguage.googleapis.com");
        port = INTERNET_DEFAULT_HTTPS_PORT;
        is_secure = TRUE;
        
        wchar_t *model_wide = NULL;
        wchar_t *key_wide = NULL;

        int model_len = MultiByteToWideChar(CP_UTF8, 0, model_utf8, -1, NULL, 0);
        if (model_len > 0) {
            model_wide = (wchar_t *)malloc(model_len * sizeof(wchar_t));
            if (model_wide) MultiByteToWideChar(CP_UTF8, 0, model_utf8, -1, model_wide, model_len);
        }
        int key_len = MultiByteToWideChar(CP_UTF8, 0, api_key_utf8, -1, NULL, 0);
        if (key_len > 0) {
            key_wide = (wchar_t *)malloc(key_len * sizeof(wchar_t));
            if (key_wide) MultiByteToWideChar(CP_UTF8, 0, api_key_utf8, -1, key_wide, key_len);
        }

        if (!model_wide || !key_wide) {
            set_error_text(out_error_utf8, "Memory allocation failed for API request parameters");
            free(model_wide);
            free(key_wide);
            goto cleanup;
        }

        path_wide = (wchar_t *)malloc(2048 * sizeof(wchar_t));
        if (path_wide) {
            swprintf(path_wide, 2048, L"/v1beta/models/%ls:generateContent?key=%ls", model_wide, key_wide);
        }
        free(model_wide);
        free(key_wide);
        
        if (!path_wide) {
            goto cleanup;
        }
    }

    connect = WinHttpConnect(session, host_wide, port, 0);
    if (!connect) {
        wchar_t err_msg[512];
        swprintf(err_msg, _countof(err_msg), L"Failed to connect to host: %ls", host_wide);
        char *err_utf8 = wide_to_utf8_alloc(err_msg);
        set_error_text(out_error_utf8, err_utf8);
        free(err_utf8);
        goto cleanup;
    }

    DWORD request_flags = is_secure ? WINHTTP_FLAG_SECURE : 0;
    request = WinHttpOpenRequest(connect, L"POST", path_wide, NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, request_flags);
    if (!request) {
        set_error_text(out_error_utf8, "Failed to open HTTP request");
        goto cleanup;
    }

    wchar_t headers[1024] = L"Content-Type: application/json\r\n";
    if (custom_url_utf8 && custom_url_utf8[0] != '\0' && api_key_utf8 && api_key_utf8[0] != '\0') {
        int key_len = MultiByteToWideChar(CP_UTF8, 0, api_key_utf8, -1, NULL, 0);
        wchar_t *key_wide = (wchar_t *)malloc(key_len * sizeof(wchar_t));
        if (key_wide) {
            MultiByteToWideChar(CP_UTF8, 0, api_key_utf8, -1, key_wide, key_len);
            swprintf(headers, _countof(headers), L"Content-Type: application/json\r\nAuthorization: Bearer %ls\r\n", key_wide);
            free(key_wide);
        }
    }

    if (!WinHttpAddRequestHeaders(request, headers, (ULONG)-1, WINHTTP_ADDREQ_FLAG_ADD)) {
        set_error_text(out_error_utf8, "Failed to add HTTP headers");
        goto cleanup;
    }

    DWORD payload_len = (DWORD)strlen(json_payload);
    if (!WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0, (LPVOID)json_payload, payload_len, payload_len, 0)) {
        DWORD err = GetLastError();
        char msg[256];
        snprintf(msg, sizeof(msg), "WinHttpSendRequest failed with error %lu", (unsigned long)err);
        set_error_text(out_error_utf8, msg);
        goto cleanup;
    }

    if (!WinHttpReceiveResponse(request, NULL)) {
        set_error_text(out_error_utf8, "Failed to receive response from API");
        goto cleanup;
    }

    DWORD status_code = 0;
    DWORD status_code_size = sizeof(status_code);
    if (WinHttpQueryHeaders(request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                            WINHTTP_HEADER_NAME_BY_INDEX, &status_code, &status_code_size, WINHTTP_NO_HEADER_INDEX)) {
        if (status_code < 200 || status_code >= 300) {
            char msg[256];
            snprintf(msg, sizeof(msg), "HTTP status code %lu received from API", (unsigned long)status_code);
            set_error_text(out_error_utf8, msg);
        }
    }

    char *response_buffer = NULL;
    size_t response_capacity = 4096;
    size_t response_len = 0;
    response_buffer = (char *)malloc(response_capacity);
    if (!response_buffer) {
        set_error_text(out_error_utf8, "Failed to allocate memory for HTTP response");
        goto cleanup;
    }

    DWORD bytes_available = 0;
    while (WinHttpQueryDataAvailable(request, &bytes_available) && bytes_available > 0) {
        if (response_len + bytes_available >= response_capacity) {
            response_capacity = (response_len + bytes_available) * 2;
            char *new_buf = (char *)realloc(response_buffer, response_capacity);
            if (!new_buf) {
                set_error_text(out_error_utf8, "Failed to reallocate response buffer");
                break;
            }
            response_buffer = new_buf;
        }

        DWORD bytes_read = 0;
        if (!WinHttpReadData(request, response_buffer + response_len, bytes_available, &bytes_read)) {
            break;
        }
        response_len += bytes_read;
        response_buffer[response_len] = '\0';
    }

    if (response_len > 0) {
        if (out_response_utf8) {
            *out_response_utf8 = response_buffer;
        } else {
            free(response_buffer);
        }
        success = (status_code >= 200 && status_code < 300);
    } else {
        free(response_buffer);
        if (status_code < 200 || status_code >= 300) {
            // Error was already set
        } else {
            set_error_text(out_error_utf8, "Empty response from API");
        }
    }

cleanup:
    free(host_wide);
    free(path_wide);
    if (request) WinHttpCloseHandle(request);
    if (connect) WinHttpCloseHandle(connect);
    if (session) WinHttpCloseHandle(session);
    return success;
}

BOOL gemini_transcribe_audio(const wchar_t *wav_path,
                             const char *api_key,
                             const char *model_name,
                             const char *custom_url,
                             const char *target_lang,
                             const char *style,
                             const char *custom_prompt,
                             const char *replace_rules,
                             const char *ai_vocab,
                             char **out_utf8_text,
                             char **out_error_utf8) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    LARGE_INTEGER file_size = {0};
    unsigned char *wav_data = NULL;
    DWORD bytes_read = 0;
    char *base64_data = NULL;
    char *prompt = NULL;
    char *escaped_prompt = NULL;
    char *payload = NULL;
    char *response_utf8 = NULL;
    BOOL success = FALSE;

    if (!wav_path || !out_utf8_text) {
        set_error_text(out_error_utf8, "Invalid arguments for audio transcriber");
        return FALSE;
    }

    *out_utf8_text = NULL;
    if (out_error_utf8) {
        free(*out_error_utf8);
        *out_error_utf8 = NULL;
    }

    file_handle = CreateFileW(wav_path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file_handle == INVALID_HANDLE_VALUE) {
        set_error_text(out_error_utf8, "Failed to open WAV file for upload");
        return FALSE;
    }

    if (!GetFileSizeEx(file_handle, &file_size) || file_size.QuadPart == 0 || file_size.QuadPart > 15 * 1024 * 1024) {
        CloseHandle(file_handle);
        set_error_text(out_error_utf8, "WAV file is invalid or exceeds 15MB limit");
        return FALSE;
    }

    wav_data = (unsigned char *)malloc((size_t)file_size.QuadPart);
    if (!wav_data) {
        CloseHandle(file_handle);
        set_error_text(out_error_utf8, "Out of memory reading WAV file data");
        return FALSE;
    }

    if (!ReadFile(file_handle, wav_data, (DWORD)file_size.QuadPart, &bytes_read, NULL) || bytes_read != (DWORD)file_size.QuadPart) {
        free(wav_data);
        CloseHandle(file_handle);
        set_error_text(out_error_utf8, "Failed to read WAV bytes into buffer");
        return FALSE;
    }
    CloseHandle(file_handle);

    size_t base64_len = 0;
    base64_data = base64_encode(wav_data, (size_t)file_size.QuadPart, &base64_len);
    free(wav_data);
    if (!base64_data) {
        set_error_text(out_error_utf8, "Base64 encoding of WAV audio failed");
        return FALSE;
    }

    prompt = build_transcribe_prompt(target_lang, style, custom_prompt, replace_rules, ai_vocab);
    if (!prompt) {
        set_error_text(out_error_utf8, "Failed to build transcription instructions prompt");
        goto cleanup;
    }

    escaped_prompt = json_escape(prompt);
    if (!escaped_prompt) {
        set_error_text(out_error_utf8, "Failed to escape transcription instructions JSON");
        goto cleanup;
    }

    size_t payload_cap = base64_len + strlen(escaped_prompt) + 1024;
    payload = (char *)malloc(payload_cap);
    if (!payload) {
        set_error_text(out_error_utf8, "Out of memory allocating multimodality payload buffer");
        goto cleanup;
    }

    sprintf_s(payload, payload_cap,
              "{\"contents\":[{\"parts\":["
              "{\"inlineData\":{\"mimeType\":\"audio/wav\",\"data\":\"%s\"}},"
              "{\"text\":\"%s\"}"
              "]}],\"generationConfig\":{\"temperature\":0.1}}",
              base64_data, escaped_prompt);

    success = gemini_send_request(api_key, model_name, custom_url, payload, &response_utf8, out_error_utf8);
    free(payload);

    if (success && response_utf8) {
        char *extracted = extract_json_text_field(response_utf8);
        free(response_utf8);
        if (extracted) {
            *out_utf8_text = extracted;
        } else {
            *out_utf8_text = _strdup("");
        }
    } else {
        if (response_utf8) {
            char *err_msg = extract_json_text_field(response_utf8);
            if (err_msg) {
                set_error_text(out_error_utf8, err_msg);
                free(err_msg);
            } else {
                set_error_text(out_error_utf8, response_utf8);
            }
            free(response_utf8);
        }
    }

cleanup:
    free(base64_data);
    free(prompt);
    free(escaped_prompt);
    return success;
}

BOOL gemini_polish_text(const char *input_text,
                        const char *api_key,
                        const char *model_name,
                        const char *custom_url,
                        int ai_engine,
                        const char *target_lang,
                        const char *style,
                        const char *custom_prompt,
                        const char *replace_rules,
                        const char *ai_vocab,
                        char **out_utf8_text,
                        char **out_error_utf8) {
    BOOL success = FALSE;
    char *payload = NULL;

    if (ai_engine == 1) {
        if (!custom_url || custom_url[0] == '\0') {
            set_error_text(out_error_utf8, "使用本地/自訂 AI 引擎時，必須配置自訂 API URL（如 http://127.0.0.1:1234）");
            return FALSE;
        }
        payload = build_openai_chat_payload(input_text, model_name, target_lang, style, custom_prompt, replace_rules, ai_vocab);
        if (!payload) {
            set_error_text(out_error_utf8, "Failed to build OpenAI chat completion payload");
            return FALSE;
        }
    } else {
        char *prompt = build_polish_prompt(input_text, target_lang, style, custom_prompt, replace_rules, ai_vocab);
        if (!prompt) {
            set_error_text(out_error_utf8, "Failed to build optimization prompt");
            return FALSE;
        }

        char *escaped_prompt = json_escape(prompt);
        free(prompt);
        if (!escaped_prompt) {
            set_error_text(out_error_utf8, "Failed to escape JSON prompt");
            return FALSE;
        }

        size_t payload_cap = strlen(escaped_prompt) + 512;
        payload = (char *)malloc(payload_cap);
        if (!payload) {
            free(escaped_prompt);
            set_error_text(out_error_utf8, "Out of memory allocating payload buffer");
            return FALSE;
        }

        sprintf_s(payload, payload_cap,
                  "{\"contents\":[{\"parts\":[{\"text\":\"%s\"}]}],\"generationConfig\":{\"temperature\":0.1}}",
                  escaped_prompt);
        free(escaped_prompt);
    }

    char *response_utf8 = NULL;
    success = gemini_send_request(api_key, model_name, custom_url, payload, &response_utf8, out_error_utf8);
    free(payload);

    if (success && response_utf8) {
        char *extracted = extract_json_text_field(response_utf8);
        free(response_utf8);
        if (extracted) {
            *out_utf8_text = extracted;
        } else {
            *out_utf8_text = _strdup("");
        }
    } else {
        if (response_utf8) {
            char *err_msg = extract_json_text_field(response_utf8);
            if (err_msg) {
                set_error_text(out_error_utf8, err_msg);
                free(err_msg);
            } else {
                set_error_text(out_error_utf8, response_utf8);
            }
            free(response_utf8);
        }
    }
    return success;
}

void asr_init_prompts(void) {
    char *p;
    
    // 1. system_prompt_openai.txt
    p = load_system_prompt(L"system_prompt_openai.txt", 
        "你是一个极其简练的语音输入法后台文本优化助手。你的任务是对用户的语音识别文本内容进行正确的加标点、去填充词（如：嗯、呃、那个）和重复句子修正。\n"
        "注意遵守规则：\n"
        "1. 你必须且只能直接输出优化后的文本，绝对不能包含任何问候、解释、分析、对话或回答文本中的内容。\n"
        "2. 英文缩写拼写修正：必须把被拆开的英文缩写合并并转为大写（如：把 'a i' 合并为 'AI'，把 'i p' 合并为 'IP'，把 'a p i' 合并为 'API'），绝对不能把 'a i' 中的字母 'a' 误判为语气词 '啊/a' 而删除！");
    free(p);

    // 2. prompt_openai_fewshot.txt
    p = load_system_prompt(L"prompt_openai_fewshot.txt", 
        "user: 那个今天天气呃挺好的吧\n"
        "assistant: 今天天气挺好的吧。\n"
        "user: 这个是结合 a i 自动纠错比那个更智能\n"
        "assistant: 这个是结合 AI 自动纠错，比那个更智能。\n"
        "user: 你是谁呀\n"
        "assistant: 你是谁呀。\n"
        "user: 不换嘛那你得当时说一下看看才知道\n"
        "assistant: 不换嘛，那你得当时说一下，看看才知道。\n");
    free(p);

    // 3. system_prompt_gemini.txt
    p = load_system_prompt(L"system_prompt_gemini.txt", 
        "你是一个专业的语音输入优化助理。请对以下【原始转写文字】进行智能优化、重写与排版。\n"
        "【核心任务】:\n"
        "1. 移除无意义的口头语、填充词（如“嗯”、“呃”、“啊”、“然后”、“那个”、“你知道的”等）。\n"
        "2. 移除不必要的重复词汇与结巴。\n"
        "3. 自动识别并处理自我修正（例如“不对，应该上午十点” -> 直接改为“上午十点”）。\n"
        "4. 自动将口述的清单、步骤或要点整理成干净、结构化的分行列表或标点符号排版。\n"
        "【输出规则】:\n"
        "- 请“仅”输出优化后的最终文本，不要包含任何解释、Markdown代码块标记、导言或括号说明。\n"
        "- 如果原始文字自动识别为无意义的噪音、空字串或无法理清的语音，请直接输出空字串。");
    free(p);

    // 4. system_prompt_gemini_transcribe.txt
    p = load_system_prompt(L"system_prompt_gemini_transcribe.txt", 
        "请将附带的语音直接转录为文字，并同时进行智能优化、重写与排版。\n"
        "【核心任务】:\n"
        "1. 自动去除口头赘词与无意义填充词（如“嗯”、“呃”、“啊”、“然后”、“那个”、“你知道的”等）。\n"
        "2. 消除语音中的结巴与重复词汇。\n"
        "3. 自动识别并处理自我修正（只保留修正后的最终意图，不要保留口误）。\n"
        "4. 自动将口述清单、步骤或重点整理为干净分行列表或合适的标点符号排版。");
    free(p);

    // 5. prompt_style_default.txt
    p = load_system_prompt(L"prompt_style_default.txt", "请优化语句，使其流畅、重点清晰。");
    free(p);

    // 6. prompt_style_business.txt
    p = load_system_prompt(L"prompt_style_business.txt", "请将文本改写得更加商务、正式、专业，适合职场与商务邮件沟通。");
    free(p);

    // 7. prompt_style_casual.txt
    p = load_system_prompt(L"prompt_style_casual.txt", "请保持日常口语风格，使语句流畅自然，不要过于死板。");
    free(p);

    // 8. prompt_style_concise.txt
    p = load_system_prompt(L"prompt_style_concise.txt", "请尽可能简洁明了，去掉赘字，保留核心重点。");
    free(p);
}

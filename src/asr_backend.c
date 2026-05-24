#include "asr_backend.h"

#include <ctype.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <winhttp.h>

#define SHERPA_OUTPUT_MAX (256 * 1024)

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

    // 3. Setup WinHTTP & WebSocket upgrade handshake
    session = WinHttpOpen(L"VoiceIME/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) {
        free(payload);
        set_error_text(out_error_utf8, "WinHttpOpen failed");
        return FALSE;
    }

    DWORD timeout = 2000;
    WinHttpSetOption(session, WINHTTP_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));

    connect = WinHttpConnect(session, L"127.0.0.1", 6006, 0);
    if (!connect) {
        free(payload);
        WinHttpCloseHandle(session);
        set_error_text(out_error_utf8, "Failed to connect to local port 6006");
        return FALSE;
    }

    request = WinHttpOpenRequest(connect, L"GET", L"/", NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
    if (!request) {
        free(payload);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        set_error_text(out_error_utf8, "WinHttpOpenRequest failed");
        return FALSE;
    }

    if (!WinHttpSetOption(request, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, NULL, 0)) {
        free(payload);
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        set_error_text(out_error_utf8, "Failed to set upgrade to websocket option");
        return FALSE;
    }

    if (!WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
        free(payload);
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
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

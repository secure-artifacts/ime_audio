#include "groq_client.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <winhttp.h>

static char *copy_text_limited(const char *text, size_t max_len) {
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
    *out_error_utf8 = copy_text_limited(message, 1500);
}

static void set_error_with_detail(char **out_error_utf8, const char *message, const char *detail) {
    size_t message_len = 0;
    size_t detail_len = 0;
    char *merged = NULL;

    if (!out_error_utf8 || !message) {
        return;
    }

    if (!detail || detail[0] == '\0') {
        set_error_text(out_error_utf8, message);
        return;
    }

    message_len = strlen(message);
    detail_len = strlen(detail);
    if (detail_len > 1100) {
        detail_len = 1100;
    }

    merged = (char *)malloc(message_len + detail_len + 4);
    if (!merged) {
        set_error_text(out_error_utf8, message);
        return;
    }

    memcpy(merged, message, message_len);
    merged[message_len] = ':';
    merged[message_len + 1] = ' ';
    memcpy(merged + message_len + 2, detail, detail_len);
    merged[message_len + 2 + detail_len] = '\0';

    free(*out_error_utf8);
    *out_error_utf8 = merged;
}

static void set_error_from_win32(char **out_error_utf8, const char *api_name) {
    DWORD error_code = GetLastError();
    char buffer[160];

    snprintf(buffer, sizeof(buffer), "%s failed with Win32 error %lu", api_name, (unsigned long)error_code);
    set_error_text(out_error_utf8, buffer);
}

static wchar_t *utf8_to_wide(const char *utf8_text) {
    int needed = 0;
    wchar_t *wide_text = NULL;

    if (!utf8_text) {
        return NULL;
    }

    needed = MultiByteToWideChar(CP_UTF8, 0, utf8_text, -1, NULL, 0);
    if (needed <= 0) {
        return NULL;
    }

    wide_text = (wchar_t *)malloc((size_t)needed * sizeof(wchar_t));
    if (!wide_text) {
        return NULL;
    }

    if (MultiByteToWideChar(CP_UTF8, 0, utf8_text, -1, wide_text, needed) <= 0) {
        free(wide_text);
        return NULL;
    }

    return wide_text;
}

static BOOL read_file_bytes(const wchar_t *path, unsigned char **out_data, DWORD *out_size) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    LARGE_INTEGER file_size = {0};
    unsigned char *buffer = NULL;
    DWORD bytes_read = 0;

    *out_data = NULL;
    *out_size = 0;

    file_handle = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file_handle == INVALID_HANDLE_VALUE) {
        return FALSE;
    }

    if (!GetFileSizeEx(file_handle, &file_size) || file_size.QuadPart <= 0 || file_size.QuadPart > MAXDWORD) {
        CloseHandle(file_handle);
        return FALSE;
    }

    buffer = (unsigned char *)malloc((size_t)file_size.QuadPart);
    if (!buffer) {
        CloseHandle(file_handle);
        return FALSE;
    }

    if (!ReadFile(file_handle, buffer, (DWORD)file_size.QuadPart, &bytes_read, NULL) || bytes_read != (DWORD)file_size.QuadPart) {
        free(buffer);
        CloseHandle(file_handle);
        return FALSE;
    }

    CloseHandle(file_handle);
    *out_data = buffer;
    *out_size = bytes_read;
    return TRUE;
}

static BOOL append_text_part(char **target, size_t *target_len, const char *text) {
    size_t text_len = strlen(text);
    char *new_buffer = (char *)realloc(*target, *target_len + text_len + 1);

    if (!new_buffer) {
        return FALSE;
    }

    *target = new_buffer;
    memcpy(*target + *target_len, text, text_len);
    *target_len += text_len;
    (*target)[*target_len] = '\0';
    return TRUE;
}

static BOOL read_response_body(HINTERNET request, char **out_body) {
    char *buffer = NULL;
    size_t total = 0;

    *out_body = NULL;

    for (;;) {
        DWORD chunk_size = 0;
        DWORD downloaded = 0;
        char *new_buffer = NULL;

        if (!WinHttpQueryDataAvailable(request, &chunk_size)) {
            free(buffer);
            return FALSE;
        }

        if (chunk_size == 0) {
            break;
        }

        new_buffer = (char *)realloc(buffer, total + chunk_size + 1);
        if (!new_buffer) {
            free(buffer);
            return FALSE;
        }
        buffer = new_buffer;

        if (!WinHttpReadData(request, buffer + total, chunk_size, &downloaded)) {
            free(buffer);
            return FALSE;
        }

        total += downloaded;
        buffer[total] = '\0';
    }

    if (!buffer) {
        buffer = (char *)malloc(1);
        if (!buffer) {
            return FALSE;
        }
        buffer[0] = '\0';
    }

    *out_body = buffer;
    return TRUE;
}

BOOL groq_transcribe_wav(const wchar_t *wav_path,
                         const char *api_key,
                         char **out_utf8_text,
                         char **out_error_utf8) {
    const char *boundary = "----VoiceImeBoundaryZ4xk9s7a";
    const char *prefix_format =
        "--%s\r\n"
        "Content-Disposition: form-data; name=\"model\"\r\n\r\n"
        "whisper-large-v3-turbo\r\n"
        "--%s\r\n"
        "Content-Disposition: form-data; name=\"language\"\r\n\r\n"
        "zh\r\n"
        "--%s\r\n"
        "Content-Disposition: form-data; name=\"response_format\"\r\n\r\n"
        "text\r\n"
        "--%s\r\n"
        "Content-Disposition: form-data; name=\"file\"; filename=\"record.wav\"\r\n"
        "Content-Type: audio/wav\r\n\r\n";
    const char *suffix_format = "\r\n--%s--\r\n";

    unsigned char *wav_data = NULL;
    DWORD wav_size = 0;
    char *prefix = NULL;
    char *suffix = NULL;
    char *body = NULL;
    size_t prefix_len = 0;
    size_t suffix_len = 0;
    size_t body_len = 0;

    wchar_t *api_key_wide = NULL;
    wchar_t *boundary_wide = NULL;
    wchar_t *headers = NULL;
    size_t header_len = 0;

    HINTERNET session = NULL;
    HINTERNET connect = NULL;
    HINTERNET request = NULL;

    DWORD status_code = 0;
    DWORD status_size = sizeof(status_code);
    char *response_body = NULL;
    BOOL ok = FALSE;

    if (!wav_path || !api_key || !out_utf8_text || api_key[0] == '\0') {
        set_error_text(out_error_utf8, "Invalid arguments or empty API key");
        return FALSE;
    }

    *out_utf8_text = NULL;
    if (out_error_utf8) {
        free(*out_error_utf8);
        *out_error_utf8 = NULL;
    }

    if (!read_file_bytes(wav_path, &wav_data, &wav_size)) {
        set_error_text(out_error_utf8, "Failed to read WAV file");
        return FALSE;
    }

    prefix_len = (size_t)snprintf(NULL, 0, prefix_format, boundary, boundary, boundary, boundary);
    suffix_len = (size_t)snprintf(NULL, 0, suffix_format, boundary);

    prefix = (char *)malloc(prefix_len + 1);
    suffix = (char *)malloc(suffix_len + 1);
    if (!prefix || !suffix) {
        set_error_text(out_error_utf8, "Out of memory while preparing multipart boundary");
        goto cleanup;
    }

    snprintf(prefix, prefix_len + 1, prefix_format, boundary, boundary, boundary, boundary);
    snprintf(suffix, suffix_len + 1, suffix_format, boundary);

    body_len = prefix_len + wav_size + suffix_len;
    if (body_len > (size_t)MAXDWORD) {
        set_error_text(out_error_utf8, "WAV file too large for WinHTTP request body");
        goto cleanup;
    }

    body = (char *)malloc(body_len);
    if (!body) {
        set_error_text(out_error_utf8, "Out of memory while preparing HTTP body");
        goto cleanup;
    }

    memcpy(body, prefix, prefix_len);
    memcpy(body + prefix_len, wav_data, wav_size);
    memcpy(body + prefix_len + wav_size, suffix, suffix_len);

    api_key_wide = utf8_to_wide(api_key);
    boundary_wide = utf8_to_wide(boundary);
    if (!api_key_wide || !boundary_wide) {
        set_error_text(out_error_utf8, "Failed to convert API key or boundary to UTF-16");
        goto cleanup;
    }

    header_len = wcslen(api_key_wide) + wcslen(boundary_wide) + 96;
    headers = (wchar_t *)malloc((header_len + 1) * sizeof(wchar_t));
    if (!headers) {
        set_error_text(out_error_utf8, "Out of memory while building request headers");
        goto cleanup;
    }

    swprintf(headers, header_len + 1,
             L"Authorization: Bearer %ls\r\nContent-Type: multipart/form-data; boundary=%ls\r\n",
             api_key_wide, boundary_wide);

    session = WinHttpOpen(L"VoiceIME/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) {
        set_error_from_win32(out_error_utf8, "WinHttpOpen");
        goto cleanup;
    }

    connect = WinHttpConnect(session, L"api.groq.com", INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!connect) {
        set_error_from_win32(out_error_utf8, "WinHttpConnect");
        goto cleanup;
    }

    request = WinHttpOpenRequest(connect, L"POST", L"/openai/v1/audio/transcriptions",
                                 NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                                 WINHTTP_FLAG_SECURE);
    if (!request) {
        set_error_from_win32(out_error_utf8, "WinHttpOpenRequest");
        goto cleanup;
    }

    if (!WinHttpSendRequest(request, headers, (DWORD)-1L, body, (DWORD)body_len, (DWORD)body_len, 0)) {
        set_error_from_win32(out_error_utf8, "WinHttpSendRequest");
        goto cleanup;
    }

    if (!WinHttpReceiveResponse(request, NULL)) {
        set_error_from_win32(out_error_utf8, "WinHttpReceiveResponse");
        goto cleanup;
    }

    if (!WinHttpQueryHeaders(request,
                             WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                             WINHTTP_HEADER_NAME_BY_INDEX,
                             &status_code,
                             &status_size,
                             WINHTTP_NO_HEADER_INDEX)) {
        set_error_from_win32(out_error_utf8, "WinHttpQueryHeaders");
        goto cleanup;
    }

    if (!read_response_body(request, &response_body)) {
        set_error_from_win32(out_error_utf8, "WinHttpReadData");
        goto cleanup;
    }

    if (status_code < 200 || status_code >= 300) {
        char status_text[128];
        snprintf(status_text, sizeof(status_text), "Groq request failed with HTTP %lu", (unsigned long)status_code);
        set_error_with_detail(out_error_utf8, status_text, response_body);
        goto cleanup;
    }

    if (response_body[0] == '\0') {
        set_error_text(out_error_utf8, "Groq returned empty transcription text");
        goto cleanup;
    }

    *out_utf8_text = response_body;
    response_body = NULL;
    ok = TRUE;

cleanup:
    if (request) {
        WinHttpCloseHandle(request);
    }
    if (connect) {
        WinHttpCloseHandle(connect);
    }
    if (session) {
        WinHttpCloseHandle(session);
    }

    free(headers);
    free(boundary_wide);
    free(api_key_wide);
    free(body);
    free(suffix);
    free(prefix);
    free(wav_data);
    free(response_body);
    return ok;
}

void groq_free_text(char *text) {
    free(text);
}

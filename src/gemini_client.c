#include "gemini_client.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
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
    // Replace newlines with spaces for single-line logging and UI display
    char *clean_msg = copy_text_limited(message, 1500);
    if (clean_msg) {
        char *p = clean_msg;
        while (*p) {
            if (*p == '\r' || *p == '\n') {
                *p = ' ';
            }
            p++;
        }
    }
    *out_error_utf8 = clean_msg;
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

static char *extract_gemini_text(const char *json) {
    const char *p = strstr(json, "\"text\"");
    if (!p) return NULL;
    p = strchr(p, ':');
    if (!p) return NULL;
    p++;
    while (*p == ' ' || *p == '\t' || *p == '\r' || *p == '\n') p++;
    if (*p != '"') return NULL;
    p++;
    
    const char *start = p;
    const char *q = start;
    while (*q) {
        if (*q == '"' && *(q - 1) != '\\') {
            break;
        }
        q++;
    }
    if (*q != '"') return NULL;
    
    size_t len = q - start;
    char *out = (char *)malloc(len + 1);
    if (!out) return NULL;
    
    size_t out_len = 0;
    for (size_t i = 0; i < len; i++) {
        if (start[i] == '\\' && i + 1 < len) {
            i++;
            if (start[i] == 'n') {
                out[out_len++] = '\n';
            } else if (start[i] == 'r') {
                out[out_len++] = '\r';
            } else if (start[i] == 't') {
                out[out_len++] = '\t';
            } else if (start[i] == '"') {
                out[out_len++] = '"';
            } else if (start[i] == '\\') {
                out[out_len++] = '\\';
            } else {
                out[out_len++] = start[i];
            }
        } else {
            out[out_len++] = start[i];
        }
    }
    out[out_len] = '\0';
    return out;
}

static char *escape_json_string(const char *input) {
    size_t len = strlen(input);
    size_t out_len = 0;
    for (size_t i = 0; i < len; i++) {
        if (input[i] == '"' || input[i] == '\\' || input[i] == '\n' || input[i] == '\r' || input[i] == '\t') {
            out_len += 2;
        } else {
            out_len += 1;
        }
    }
    char *out = (char *)malloc(out_len + 1);
    if (!out) return NULL;
    
    size_t j = 0;
    for (size_t i = 0; i < len; i++) {
        if (input[i] == '"') { out[j++] = '\\'; out[j++] = '"'; }
        else if (input[i] == '\\') { out[j++] = '\\'; out[j++] = '\\'; }
        else if (input[i] == '\n') { out[j++] = '\\'; out[j++] = 'n'; }
        else if (input[i] == '\r') { out[j++] = '\\'; out[j++] = 'r'; }
        else if (input[i] == '\t') { out[j++] = '\\'; out[j++] = 't'; }
        else { out[j++] = input[i]; }
    }
    out[j] = '\0';
    return out;
}

static BOOL do_gemini_request(const char *api_key,
                              const char *project_id,
                              const char *model_id,
                              const char *json_body,
                              char **out_text_utf8,
                              char **out_error_utf8) {
    BOOL ok = FALSE;
    HINTERNET session = NULL;
    HINTERNET connect = NULL;
    HINTERNET request = NULL;
    DWORD status_code = 0;
    DWORD status_size = sizeof(status_code);
    char *response_body = NULL;
    
    wchar_t *host_name = L"generativelanguage.googleapis.com";
    wchar_t path[1024];
    wchar_t *model_id_wide = utf8_to_wide(model_id);
    wchar_t *api_key_wide = utf8_to_wide(api_key);
    
    swprintf(path, _countof(path), L"/v1beta/models/%ls:generateContent?key=%ls", model_id_wide, api_key_wide);

    session = WinHttpOpen(L"VoiceIME/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) { set_error_from_win32(out_error_utf8, "WinHttpOpen"); goto cleanup; }

    connect = WinHttpConnect(session, host_name, INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!connect) { set_error_from_win32(out_error_utf8, "WinHttpConnect"); goto cleanup; }

    request = WinHttpOpenRequest(connect, L"POST", path,
                                 NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                                 WINHTTP_FLAG_SECURE);
    if (!request) { set_error_from_win32(out_error_utf8, "WinHttpOpenRequest"); goto cleanup; }

    wchar_t *headers = L"Content-Type: application/json\r\n";
    if (!WinHttpSendRequest(request, headers, (DWORD)-1L, (LPVOID)json_body, (DWORD)strlen(json_body), (DWORD)strlen(json_body), 0)) {
        set_error_from_win32(out_error_utf8, "WinHttpSendRequest"); goto cleanup;
    }

    if (!WinHttpReceiveResponse(request, NULL)) {
        set_error_from_win32(out_error_utf8, "WinHttpReceiveResponse"); goto cleanup;
    }

    if (!WinHttpQueryHeaders(request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                             WINHTTP_HEADER_NAME_BY_INDEX, &status_code, &status_size, WINHTTP_NO_HEADER_INDEX)) {
        set_error_from_win32(out_error_utf8, "WinHttpQueryHeaders"); goto cleanup;
    }

    if (!read_response_body(request, &response_body)) {
        set_error_from_win32(out_error_utf8, "WinHttpReadData"); goto cleanup;
    }

    if (status_code < 200 || status_code >= 300) {
        char status_text[1024];
        snprintf(status_text, sizeof(status_text), "Gemini request failed with HTTP %lu: %s", (unsigned long)status_code, response_body);
        set_error_text(out_error_utf8, status_text);
        goto cleanup;
    }

    char *extracted = extract_gemini_text(response_body);
    if (!extracted) {
        set_error_text(out_error_utf8, "Failed to parse Gemini response");
        goto cleanup;
    }

    *out_text_utf8 = extracted;
    ok = TRUE;

cleanup:
    if (request) WinHttpCloseHandle(request);
    if (connect) WinHttpCloseHandle(connect);
    if (session) WinHttpCloseHandle(session);
    free(api_key_wide);
    free(model_id_wide);
    free(response_body);
    return ok;
}

BOOL gemini_process_text(const char *api_key,
                         const char *project_id,
                         const char *model_id,
                         const char *custom_prompt,
                         const char *target_lang,
                         const char *thinking_level,
                         const char *input_text,
                         char **out_text_utf8,
                         char **out_error_utf8) {
    if (!api_key || !model_id || !input_text || !out_text_utf8) {
        set_error_text(out_error_utf8, "Invalid arguments");
        return FALSE;
    }

    *out_text_utf8 = NULL;
    if (out_error_utf8 && *out_error_utf8) {
        free(*out_error_utf8);
        *out_error_utf8 = NULL;
    }

    const char *sys_prompt = (custom_prompt && custom_prompt[0] != '\0') ? custom_prompt : "You are an AI text processing assistant.";

    char prompt_format_trans[] = "Please translate the following text into %s. Output ONLY the translated text, without any conversational filler, markdown formatting, or quotes.\\n\\n%s";
    char prompt_format_refine[] = "Please process the following text according to the system instructions. Output ONLY the processed text, without any conversational filler, markdown formatting, or quotes.\\n\\n%s";

    char *escaped_input = escape_json_string(input_text);
    if (!escaped_input) {
        set_error_text(out_error_utf8, "Failed to escape input text");
        return FALSE;
    }

    char *escaped_sys_prompt = escape_json_string(sys_prompt);
    if (!escaped_sys_prompt) {
        free(escaped_input);
        set_error_text(out_error_utf8, "Failed to escape system prompt");
        return FALSE;
    }

    char *prompt = NULL;
    if (target_lang && target_lang[0] != '\0') {
        size_t prompt_len = strlen(prompt_format_trans) + strlen(target_lang) + strlen(escaped_input) + 1;
        prompt = (char *)malloc(prompt_len);
        snprintf(prompt, prompt_len, prompt_format_trans, target_lang, escaped_input);
    } else {
        size_t prompt_len = strlen(prompt_format_refine) + strlen(escaped_input) + 1;
        prompt = (char *)malloc(prompt_len);
        snprintf(prompt, prompt_len, prompt_format_refine, escaped_input);
    }
    free(escaped_input);

    char thinking_json[256] = "";
    if (thinking_level && thinking_level[0] != '\0' && strcmp(thinking_level, "NONE") != 0) {
        snprintf(thinking_json, sizeof(thinking_json), "\"thinkingConfig\":{\"thinkingLevel\":\"%s\"},", thinking_level);
    }

    char json_format[] = "{"
        "\"systemInstruction\":{\"parts\":[{\"text\":\"%s\"}]},"
        "\"contents\":[{\"parts\":[{\"text\":\"%s\"}]}],"
        "\"generationConfig\":{"
            "\"temperature\":1.0,"
            "\"topP\":0.95,"
            "%s"
            "\"maxOutputTokens\":65535"
        "},"
        "\"safetySettings\":["
            "{\"category\":\"HARM_CATEGORY_HATE_SPEECH\",\"threshold\":\"OFF\"},"
            "{\"category\":\"HARM_CATEGORY_DANGEROUS_CONTENT\",\"threshold\":\"OFF\"},"
            "{\"category\":\"HARM_CATEGORY_SEXUALLY_EXPLICIT\",\"threshold\":\"OFF\"},"
            "{\"category\":\"HARM_CATEGORY_HARASSMENT\",\"threshold\":\"OFF\"}"
        "]"
    "}";
    size_t json_len = strlen(json_format) + strlen(escaped_sys_prompt) + strlen(prompt) + strlen(thinking_json) + 1;
    char *json_body = (char *)malloc(json_len);
    snprintf(json_body, json_len, json_format, escaped_sys_prompt, prompt, thinking_json);
    free(escaped_sys_prompt);
    free(prompt);

    BOOL res = do_gemini_request(api_key, project_id, model_id, json_body, out_text_utf8, out_error_utf8);
    free(json_body);

    return res;
}

void gemini_free_text(char *text) {
    free(text);
}

static BOOL read_file_bytes_gem(const wchar_t *path, unsigned char **out_data, DWORD *out_size) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    LARGE_INTEGER file_size = {0};
    unsigned char *buffer = NULL;
    DWORD bytes_read = 0;

    *out_data = NULL;
    *out_size = 0;

    file_handle = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file_handle == INVALID_HANDLE_VALUE) return FALSE;

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

static const char base64_chars[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

static char *base64_encode(const unsigned char *data, size_t len) {
    size_t out_len = 4 * ((len + 2) / 3);
    char *out = (char *)malloc(out_len + 1);
    if (!out) return NULL;
    
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
        out[out_len - 1 - k] = '=';
    }
    out[out_len] = '\0';
    return out;
}

BOOL gemini_transcribe_wav(const wchar_t *wav_path,
                           const char *api_key,
                           const char *project_id,
                           const char *model_id,
                           const char *custom_prompt,
                           const char *target_lang,
                           const char *thinking_level,
                           char **out_utf8_text,
                           char **out_error_utf8) {
    unsigned char *wav_data = NULL;
    DWORD wav_size = 0;
    char *base64_wav = NULL;
    char *json_body = NULL;
    BOOL res = FALSE;

    if (!wav_path || !api_key || !model_id || !out_utf8_text) {
        set_error_text(out_error_utf8, "Invalid arguments");
        return FALSE;
    }
    *out_utf8_text = NULL;

    if (!read_file_bytes_gem(wav_path, &wav_data, &wav_size)) {
        set_error_text(out_error_utf8, "Failed to read WAV file");
        return FALSE;
    }

    base64_wav = base64_encode(wav_data, wav_size);
    free(wav_data);
    if (!base64_wav) {
        set_error_text(out_error_utf8, "Out of memory during base64 encoding");
        return FALSE;
    }

    char prompt_buf[4096];
    prompt_buf[0] = '\0';
    
    if (target_lang && target_lang[0] != '\0') {
        snprintf(prompt_buf, sizeof(prompt_buf), "Please translate the following audio into %s. Output ONLY the translated text, without any conversational filler, markdown, or timestamps. If the audio is silent, unclear, or contains no speech, output nothing (an empty string).", target_lang);
    } else if (custom_prompt && custom_prompt[0] != '\0') {
        snprintf(prompt_buf, sizeof(prompt_buf), "Please process the following audio according to these instructions: %s. Output ONLY the final processed text, without any conversational filler, markdown, or timestamps. If the audio is silent, unclear, or contains no speech, output nothing (an empty string).", custom_prompt);
    } else {
        snprintf(prompt_buf, sizeof(prompt_buf), "Please accurately transcribe the following audio into text. If the language is not specified, auto-detect it. Do not translate unless explicitly requested. Output ONLY the raw transcribed text, without any conversational filler, markdown, or timestamps. If the audio is silent, unclear, or contains no speech, output nothing (an empty string).");
    }

    char *escaped_prompt = escape_json_string(prompt_buf);

    char thinking_json[256] = "";
    if (thinking_level && thinking_level[0] != '\0' && strcmp(thinking_level, "NONE") != 0) {
        snprintf(thinking_json, sizeof(thinking_json), "\"thinkingConfig\":{\"thinkingLevel\":\"%s\"},", thinking_level);
    }

    char json_format[] = "{"
        "\"contents\":[{\"parts\":[{\"text\":\"%s\"},{\"inline_data\":{\"mime_type\":\"audio/wav\",\"data\":\"%s\"}}]}],"
        "\"generationConfig\":{"
            "\"temperature\":1.0,"
            "\"topP\":0.95,"
            "%s"
            "\"maxOutputTokens\":65535"
        "},"
        "\"safetySettings\":["
            "{\"category\":\"HARM_CATEGORY_HATE_SPEECH\",\"threshold\":\"OFF\"},"
            "{\"category\":\"HARM_CATEGORY_DANGEROUS_CONTENT\",\"threshold\":\"OFF\"},"
            "{\"category\":\"HARM_CATEGORY_SEXUALLY_EXPLICIT\",\"threshold\":\"OFF\"},"
            "{\"category\":\"HARM_CATEGORY_HARASSMENT\",\"threshold\":\"OFF\"}"
        "]"
    "}";
    
    size_t json_len = strlen(json_format) + strlen(escaped_prompt) + strlen(base64_wav) + strlen(thinking_json) + 1;
    json_body = (char *)malloc(json_len);
    if (!json_body) {
        set_error_text(out_error_utf8, "Out of memory preparing JSON body");
        free(escaped_prompt);
        free(base64_wav);
        return FALSE;
    }

    snprintf(json_body, json_len, json_format, escaped_prompt, base64_wav, thinking_json);
    free(escaped_prompt);
    free(base64_wav);

    res = do_gemini_request(api_key, project_id, model_id, json_body, out_utf8_text, out_error_utf8);
    free(json_body);

    return res;
}

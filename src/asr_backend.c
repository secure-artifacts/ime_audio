#include "asr_backend.h"

#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

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

    return ASR_BACKEND_GROQ;
}

const wchar_t *asr_backend_name(AsrBackendKind backend) {
    if (backend == ASR_BACKEND_SHERPA) {
        return L"sherpa";
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

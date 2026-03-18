#include <stdio.h>
#include <stdlib.h>
#include <stdarg.h>
#include <string.h>
#include <wchar.h>
#include <wctype.h>
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>

#include "asr_backend.h"
#include "audio_recorder.h"
#include "groq_client.h"
#include "gemini_client.h"
#include "input_injector.h"

#define MAIN_CLASS_NAME L"VoiceImeMainWindow"
#define FLOAT_CLASS_NAME L"VoiceImeFloatWindow"
#define APP_TITLE L"语音输入助手"

#define HOTKEY_RECORD_ID 1

#define WMAPP_TRAYICON (WM_APP + 1)
#define WMAPP_TRANSCRIBE_DONE (WM_APP + 2)
#define WMAPP_FLOAT_TOGGLE (WM_APP + 3)

#define TIMER_FOLLOW_INPUT 1
#define TRAY_ICON_ID 1001

#define IDC_EDIT_API 2001
#define IDC_EDIT_HOTKEY 2002
#define IDC_CHECK_CTRL 2003
#define IDC_CHECK_ALT 2004
#define IDC_CHECK_SHIFT 2005
#define IDC_CHECK_WIN 2006
#define IDC_BTN_APPLY 2007
#define IDC_LABEL_STATUS 2008
#define IDC_LABEL_CURRENT_HOTKEY 2009
#define IDC_COMBO_BACKEND 2010
#define IDC_CHECK_CONTINUOUS 2011
#define IDC_EDIT_THRESHOLD 2012
#define IDC_EDIT_SILENCE 2013
#define IDC_EDIT_MINREC 2014
#define IDC_EDIT_MAXREC 2015
#define IDC_EDIT_SHERPA_EXE 2016
#define IDC_EDIT_SHERPA_ARGS 2017
#define IDC_BTN_SELF_CHECK 2018
#define IDC_BTN_INSTALL_SHERPA 2019
#define IDC_LABEL_SELFCHECK 2020
#define IDC_BTN_EXIT 2021
#define IDC_CHECK_AUTO_STOP 2022
#define IDC_EDIT_REPLACE_RULES 2023
#define IDC_COMBO_MIC 2024
#define IDC_EDIT_GEMINI_KEY 2025
#define IDC_EDIT_GLADIA_KEY 2026
#define IDC_CHECK_TRANSLATE 2027
#define IDC_COMBO_LANG 2028
#define IDC_EDIT_PROJECT_ID 2029
#define IDC_COMBO_GEMINI_MODEL 2030
#define IDC_EDIT_GEMINI_PROMPT 2031
#define IDC_COMBO_THINKING 2032
#define IDC_LIST_LOG 2033
#define IDC_COMBO_MODEL 2041
#define IDC_BTN_APPLY_MODEL 2042
#define IDC_BTN_OPEN_CONFIG 2043

#define IDC_FLOAT_TOGGLE 3001
#define IDC_FLOAT_STATUS 3002

#define ID_TRAY_OPEN 4001
#define ID_TRAY_EXIT 4002

typedef enum VoiceState {
    VOICE_IDLE = 0,
    VOICE_RECORDING = 1,
    VOICE_TRANSCRIBING = 2,
    VOICE_PAUSED = 3
} VoiceState;

typedef struct AppState {
    HINSTANCE instance;
    HWND main_hwnd;
    HWND float_hwnd;
    HWND api_edit;
    HWND gemini_key_edit;
    HWND project_id_edit;
    HWND gemini_model_combo;
    HWND gemini_prompt_edit;
    HWND thinking_combo;
    HWND model_combo;
    HWND gladia_key_edit;
    HWND hotkey_edit;
    HWND check_ctrl;
    HWND check_alt;
    HWND check_shift;
    HWND check_win;
    HWND backend_combo;
    HWND mic_combo;
    HWND lang_combo;
    HWND continuous_check;
    HWND auto_stop_check;
    HWND translate_check;
    HWND threshold_edit;
    HWND silence_edit;
    HWND minrec_edit;
    HWND maxrec_edit;
    HWND sherpa_exe_edit;
    HWND sherpa_args_edit;
    HWND replace_rules_edit;
    HWND selfcheck_label;
    HWND status_label;
    HWND current_hotkey_label;
    HWND float_button;
    HWND float_status;
    HWND log_list;

    wchar_t config_path[MAX_PATH];
    wchar_t wav_path[MAX_PATH];
    wchar_t log_path[MAX_PATH];
    wchar_t sherpa_exe[2048];
    wchar_t sherpa_args[4096];
    wchar_t replace_rules[2048];
    wchar_t gemini_key[256];
    wchar_t project_id[256];
    wchar_t gemini_model[256];
    wchar_t gemini_prompt[1024];
    wchar_t thinking_level[32];
    wchar_t gladia_key[256];
    wchar_t target_lang[64];

    UINT hotkey_mods;
    UINT hotkey_vk;
    UINT mic_device_id;
    BOOL hotkey_registered;
    BOOL tray_added;
    BOOL exit_requested;
    BOOL continuous_mode;
    BOOL pause_requested;
    BOOL auto_stop_enabled;
    BOOL stop_after_current;
    BOOL translate_enabled;
    BOOL current_had_voice;
    int local_model_index;

    AsrBackendKind backend;
    AudioRecorderConfig recorder_config;

    VoiceState state;
    HANDLE worker_thread;

    HWND target_window;
    HWND follow_target_window;
} AppState;

typedef struct TranscribeTask {
    HWND notify_hwnd;
    wchar_t wav_path[MAX_PATH];
    wchar_t sherpa_exe[2048];
    wchar_t sherpa_args[4096];
    AsrBackendKind backend;
    char *api_key;
    char *project_id;
    char *model_id;
    HWND target_window;
    BOOL translate_enabled;
    char *custom_prompt;
    char *target_lang;
    char *thinking_level;
} TranscribeTask;

typedef struct TranscribeResult {
    BOOL success;
    char *text;
    char *error_text;
    HWND target_window;
} TranscribeResult;

static void trim_wide_whitespace(wchar_t *text);
static void add_ui_log(AppState *app, const wchar_t *text);
static void save_settings(AppState *app);
static char *wide_to_utf8_alloc(const wchar_t *wide_text);
static wchar_t *utf8_to_wide_alloc(const char *utf8_text);

static BOOL build_temp_wav_path(wchar_t *out_path, DWORD out_path_size) {
    wchar_t temp_dir[MAX_PATH];
    DWORD len = GetTempPathW(MAX_PATH, temp_dir);

    if (len == 0 || len >= MAX_PATH) {
        return FALSE;
    }

    if (swprintf(out_path, out_path_size, L"%lsvoice_ime_record.wav", temp_dir) < 0) {
        return FALSE;
    }

    return TRUE;
}

static void extract_parent_dir(const wchar_t *path, wchar_t *out_dir, size_t out_len);

static BOOL build_config_path(wchar_t *out_path, DWORD out_path_size) {
    wchar_t exe_path[MAX_PATH];
    wchar_t app_dir[MAX_PATH];
    GetModuleFileNameW(NULL, exe_path, MAX_PATH);
    extract_parent_dir(exe_path, app_dir, _countof(app_dir));
    if (swprintf(out_path, out_path_size, L"%ls\\voice_ime.ini", app_dir) > 0) {
        return TRUE;
    }
    return FALSE;
}

static BOOL build_log_path(const wchar_t *config_path, wchar_t *out_path, DWORD out_path_size) {
    wchar_t temp[MAX_PATH];
    size_t len = 0;

    if (!config_path || !out_path || out_path_size == 0) {
        return FALSE;
    }

    wcsncpy_s(temp, _countof(temp), config_path, _TRUNCATE);
    len = wcslen(temp);
    while (len > 0) {
        if (temp[len - 1] == L'\\' || temp[len - 1] == L'/') {
            break;
        }
        len--;
    }

    if (len == 0) {
        return FALSE;
    }

    temp[len] = L'\0';
    if (swprintf(out_path, out_path_size, L"%lsvoice_ime.log", temp) < 0) {
        return FALSE;
    }

    return TRUE;
}

static void extract_parent_dir(const wchar_t *path, wchar_t *out_dir, size_t out_len) {
    size_t len = 0;

    if (!path || !out_dir || out_len == 0) {
        return;
    }

    wcsncpy_s(out_dir, out_len, path, _TRUNCATE);
    len = wcslen(out_dir);
    while (len > 0) {
        if (out_dir[len - 1] == L'\\' || out_dir[len - 1] == L'/') {
            out_dir[len - 1] = L'\0';
            return;
        }
        len--;
    }

    out_dir[0] = L'\0';
}

static BOOL file_exists_non_dir(const wchar_t *path) {
    DWORD attrs;

    if (!path || path[0] == L'\0') {
        return FALSE;
    }

    attrs = GetFileAttributesW(path);
    if (attrs == INVALID_FILE_ATTRIBUTES) {
        return FALSE;
    }

    return (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

static BOOL is_absolute_path_w(const wchar_t *path) {
    if (!path || path[0] == L'\0') {
        return FALSE;
    }

    if ((wcslen(path) >= 2 && path[1] == L':') ||
        (path[0] == L'\\' && path[1] == L'\\') ||
        path[0] == L'/' || path[0] == L'\\') {
        return TRUE;
    }

    return FALSE;
}

static void resolve_path_from_base(const wchar_t *base_dir,
                                   const wchar_t *maybe_relative,
                                   wchar_t *out_path,
                                   size_t out_len) {
    if (!out_path || out_len == 0) {
        return;
    }

    out_path[0] = L'\0';
    if (!maybe_relative || maybe_relative[0] == L'\0') {
        return;
    }

    if (is_absolute_path_w(maybe_relative) || !base_dir || base_dir[0] == L'\0') {
        wcsncpy_s(out_path, out_len, maybe_relative, _TRUNCATE);
        return;
    }

    swprintf(out_path, out_len, L"%ls\\%ls", base_dir, maybe_relative);
}

static BOOL extract_option_value(const wchar_t *args,
                                 const wchar_t *prefix,
                                 wchar_t *out_value,
                                 size_t out_len) {
    size_t prefix_len = 0;
    const wchar_t *p = args;

    if (!args || !prefix || !out_value || out_len == 0) {
        return FALSE;
    }

    out_value[0] = L'\0';
    prefix_len = wcslen(prefix);

    while (*p) {
        size_t copied = 0;
        while (*p && iswspace(*p)) {
            p++;
        }

        if (_wcsnicmp(p, prefix, prefix_len) == 0) {
            const wchar_t *v = p + prefix_len;
            if (*v == L'"') {
                v++;
                while (v[copied] && v[copied] != L'"' && copied + 1 < out_len) {
                    out_value[copied] = v[copied];
                    copied++;
                }
            } else {
                while (v[copied] && !iswspace(v[copied]) && copied + 1 < out_len) {
                    out_value[copied] = v[copied];
                    copied++;
                }
            }

            out_value[copied] = L'\0';
            trim_wide_whitespace(out_value);
            return out_value[0] != L'\0';
        }

        while (*p && !iswspace(*p)) {
            p++;
        }
    }

    return FALSE;
}

static void app_log_line(const AppState *app, const char *fmt, ...) {
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    SYSTEMTIME st;
    char line[1536];
    char payload[1300];
    DWORD written = 0;
    va_list args;

    if (!app || !app->log_path[0] || !fmt) {
        return;
    }

    GetLocalTime(&st);

    va_start(args, fmt);
    vsnprintf(payload, sizeof(payload), fmt, args);
    va_end(args);

    snprintf(line,
             sizeof(line),
             "%04u-%02u-%02u %02u:%02u:%02u.%03u | %s\\r\\n",
             (unsigned)st.wYear,
             (unsigned)st.wMonth,
             (unsigned)st.wDay,
             (unsigned)st.wHour,
             (unsigned)st.wMinute,
             (unsigned)st.wSecond,
             (unsigned)st.wMilliseconds,
             payload);

    file_handle = CreateFileW(app->log_path,
                              FILE_APPEND_DATA,
                              FILE_SHARE_READ | FILE_SHARE_WRITE,
                              NULL,
                              OPEN_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL,
                              NULL);
    if (file_handle == INVALID_HANDLE_VALUE) {
        return;
    }

    WriteFile(file_handle, line, (DWORD)strlen(line), &written, NULL);
    CloseHandle(file_handle);
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

static void trim_wide_whitespace(wchar_t *text) {
    wchar_t *start = text;
    wchar_t *end = NULL;

    if (!text) {
        return;
    }

    // 同时清理空格、制表符以及换行符
    while (*start != L'\0' && (iswspace(*start) || *start == L'\r' || *start == L'\n')) {
        start++;
    }

    if (start != text) {
        memmove(text, start, (wcslen(start) + 1) * sizeof(wchar_t));
    }

    end = text + wcslen(text);
    while (end > text && (iswspace(end[-1]) || end[-1] == L'\r' || end[-1] == L'\n')) {
        end--;
    }
    *end = L'\0';
}

static BOOL contains_non_ascii_utf8(const char *text) {
    const unsigned char *p = (const unsigned char *)text;

    if (!text) {
        return FALSE;
    }

    while (*p) {
        if (*p >= 0x80) {
            return TRUE;
        }
        p++;
    }

    return FALSE;
}

static BOOL ends_with_token_utf8(const char *text, size_t text_len, const char *token) {
    size_t token_len = 0;

    if (!text || !token) {
        return FALSE;
    }

    token_len = strlen(token);
    if (token_len == 0 || token_len > text_len) {
        return FALSE;
    }

    return memcmp(text + text_len - token_len, token, token_len) == 0;
}

static BOOL contains_token_utf8(const char *text, const char *token) {
    if (!text || !token || token[0] == '\0') {
        return FALSE;
    }

    return strstr(text, token) != NULL;
}

static BOOL text_ends_with_sentence_punctuation(const char *text) {
    size_t len = 0;
    unsigned char last = 0;

    if (!text) {
        return FALSE;
    }

    len = strlen(text);
    while (len > 0 && isspace((unsigned char)text[len - 1])) {
        len--;
    }

    if (len == 0) {
        return FALSE;
    }

    last = (unsigned char)text[len - 1];
    if (last == '.' || last == '!' || last == '?' || last == ',' || last == ';' || last == ':') {
        return TRUE;
    }

    if (ends_with_token_utf8(text, len, "。") ||
        ends_with_token_utf8(text, len, "！") ||
        ends_with_token_utf8(text, len, "？") ||
        ends_with_token_utf8(text, len, "，") ||
        ends_with_token_utf8(text, len, "；") ||
        ends_with_token_utf8(text, len, "：")) {
        return TRUE;
    }

    return FALSE;
}

static const char *pick_auto_punctuation_suffix(const char *text) {
    static const char *question_suffixes[] = {"吗", "么", "嘛", "呢"};
    static const char *question_keywords[] = {
        "什么", "怎么", "为什么", "为啥", "是否", "是不是", "能不能", "可不可以",
        "要不要", "有没有", "行不行", "哪儿", "哪里", "谁", "多久", "多少", "几点", "几号"
    };
    static const char *exclaim_suffixes[] = {"啊", "呀", "哇", "啦"};
    static const char *exclaim_keywords[] = {"太好了", "太棒了", "真棒", "好厉害", "牛啊"};
    size_t len = 0;
    BOOL has_non_ascii = FALSE;
    size_t i = 0;

    if (!text || text[0] == '\0') {
        return ".";
    }

    has_non_ascii = contains_non_ascii_utf8(text);
    len = strlen(text);
    while (len > 0 && isspace((unsigned char)text[len - 1])) {
        len--;
    }

    if (len == 0) {
        return has_non_ascii ? "。" : ".";
    }

    for (i = 0; i < _countof(question_suffixes); ++i) {
        if (ends_with_token_utf8(text, len, question_suffixes[i])) {
            return has_non_ascii ? "？" : "?";
        }
    }

    for (i = 0; i < _countof(question_keywords); ++i) {
        if (contains_token_utf8(text, question_keywords[i])) {
            return has_non_ascii ? "？" : "?";
        }
    }

    for (i = 0; i < _countof(exclaim_suffixes); ++i) {
        if (ends_with_token_utf8(text, len, exclaim_suffixes[i])) {
            return has_non_ascii ? "！" : "!";
        }
    }

    for (i = 0; i < _countof(exclaim_keywords); ++i) {
        if (contains_token_utf8(text, exclaim_keywords[i])) {
            return has_non_ascii ? "！" : "!";
        }
    }

    return has_non_ascii ? "。" : ".";
}

static void append_utf8_suffix(char **text_ptr, const char *suffix) {
    size_t text_len = 0;
    size_t suffix_len = 0;
    char *merged = NULL;

    if (!text_ptr || !*text_ptr || !suffix || suffix[0] == '\0') {
        return;
    }

    text_len = strlen(*text_ptr);
    suffix_len = strlen(suffix);
    merged = (char *)realloc(*text_ptr, text_len + suffix_len + 1);
    if (!merged) {
        return;
    }

    memcpy(merged + text_len, suffix, suffix_len + 1);
    *text_ptr = merged;
}

static void apply_simple_sherpa_punctuation(AppState *app, char **text_ptr) {
    const char *suffix = NULL;

    if (!app || !text_ptr || !*text_ptr) {
        return;
    }

    if (app->backend != ASR_BACKEND_SHERPA) {
        return;
    }

    trim_ascii_whitespace(*text_ptr);
    if ((*text_ptr)[0] == '\0' || text_ends_with_sentence_punctuation(*text_ptr)) {
        return;
    }

    suffix = pick_auto_punctuation_suffix(*text_ptr);
    append_utf8_suffix(text_ptr, suffix);
    app_log_line(app, "auto punctuation appended for sherpa text suffix=%s", suffix);
}

static int replace_all_wide(wchar_t **text_ptr, const wchar_t *from, const wchar_t *to) {
    wchar_t *text = NULL;
    const wchar_t *scan = NULL;
    const wchar_t *hit = NULL;
    wchar_t *new_text = NULL;
    wchar_t *dst = NULL;
    size_t from_len = 0;
    size_t to_len = 0;
    size_t old_len = 0;
    size_t hit_count = 0;
    size_t new_len = 0;

    if (!text_ptr || !*text_ptr || !from || !to) {
        return 0;
    }

    from_len = wcslen(from);
    to_len = wcslen(to);
    if (from_len == 0) {
        return 0;
    }

    text = *text_ptr;
    old_len = wcslen(text);
    scan = text;
    while ((hit = wcsstr(scan, from)) != NULL) {
        hit_count++;
        scan = hit + from_len;
    }

    if (hit_count == 0) {
        return 0;
    }

    if (to_len >= from_len) {
        size_t grow = to_len - from_len;
        if (grow > 0 && hit_count > (((size_t)-1) - old_len - 1) / grow) {
            return 0;
        }
        new_len = old_len + hit_count * grow;
    } else {
        new_len = old_len - hit_count * (from_len - to_len);
    }

    new_text = (wchar_t *)malloc((new_len + 1) * sizeof(wchar_t));
    if (!new_text) {
        return 0;
    }

    scan = text;
    dst = new_text;
    while ((hit = wcsstr(scan, from)) != NULL) {
        size_t prefix_len = (size_t)(hit - scan);
        if (prefix_len > 0) {
            memcpy(dst, scan, prefix_len * sizeof(wchar_t));
            dst += prefix_len;
        }

        if (to_len > 0) {
            memcpy(dst, to, to_len * sizeof(wchar_t));
            dst += to_len;
        }

        scan = hit + from_len;
    }

    wcscpy_s(dst, (size_t)(new_len - (size_t)(dst - new_text) + 1), scan);

    free(*text_ptr);
    *text_ptr = new_text;
    return (int)hit_count;
}

static void apply_user_replace_rules(AppState *app, char **text_ptr) {
    wchar_t *text_wide = NULL;
    wchar_t *rules_copy = NULL;
    wchar_t *context = NULL;
    wchar_t *rule = NULL;
    int replaced_total = 0;

    if (!app || !text_ptr || !*text_ptr || app->replace_rules[0] == L'\0') {
        return;
    }

    text_wide = utf8_to_wide_alloc(*text_ptr);
    if (!text_wide) {
        return;
    }

    rules_copy = _wcsdup(app->replace_rules);
    if (!rules_copy) {
        free(text_wide);
        return;
    }

    rule = wcstok_s(rules_copy, L";\uFF1B\r\n", &context);
    while (rule) {
        wchar_t *separator = NULL;
        wchar_t *from = NULL;
        wchar_t *to = NULL;
        int sep_len = 0;

        trim_wide_whitespace(rule);
        if (rule[0] == L'\0') {
            rule = wcstok_s(NULL, L";\uFF1B\r\n", &context);
            continue;
        }

        separator = wcsstr(rule, L"=>");
        if (separator) {
            sep_len = 2;
        } else {
            separator = wcsstr(rule, L"->");
            if (separator) {
                sep_len = 2;
            } else {
                separator = wcschr(rule, L'=');
                if (separator) {
                    sep_len = 1;
                }
            }
        }

        if (!separator) {
            rule = wcstok_s(NULL, L";\uFF1B\r\n", &context);
            continue;
        }

        *separator = L'\0';
        from = rule;
        to = separator + sep_len;
        trim_wide_whitespace(from);
        trim_wide_whitespace(to);

        if (from[0] == L'\0' || wcscmp(from, to) == 0) {
            rule = wcstok_s(NULL, L";\uFF1B\r\n", &context);
            continue;
        }

        replaced_total += replace_all_wide(&text_wide, from, to);
        rule = wcstok_s(NULL, L";\uFF1B\r\n", &context);
    }

    if (replaced_total > 0) {
        char *converted = NULL;
        trim_wide_whitespace(text_wide);
        converted = wide_to_utf8_alloc(text_wide);
        if (converted) {
            free(*text_ptr);
            *text_ptr = converted;
            app_log_line(app, "replace rules applied count=%d", replaced_total);
        }
    }

    free(rules_copy);
    free(text_wide);
}

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

static wchar_t *utf8_to_wide_alloc(const char *utf8_text) {
    int needed = 0;
    wchar_t *wide = NULL;

    if (!utf8_text) {
        return NULL;
    }

    needed = MultiByteToWideChar(CP_UTF8, 0, utf8_text, -1, NULL, 0);
    if (needed <= 0) {
        return NULL;
    }

    wide = (wchar_t *)malloc((size_t)needed * sizeof(wchar_t));
    if (!wide) {
        return NULL;
    }

    if (MultiByteToWideChar(CP_UTF8, 0, utf8_text, -1, wide, needed) <= 0) {
        free(wide);
        return NULL;
    }

    return wide;
}


static void add_ui_log(AppState *app, const wchar_t *text) {
    if (!app || !app->log_list || !text) return;
    
    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t line[2048];
    swprintf(line, _countof(line), L"[%02u:%02u:%02u] %ls", st.wHour, st.wMinute, st.wSecond, text);

    LRESULT count = SendMessageW(app->log_list, LB_ADDSTRING, 0, (LPARAM)line);
    if (count != LB_ERR) {
        SendMessageW(app->log_list, LB_SETTOPINDEX, count, 0);
    }
    
    count = SendMessageW(app->log_list, LB_GETCOUNT, 0, 0);
    while (count > 100) {
        SendMessageW(app->log_list, LB_DELETESTRING, 0, 0);
        count--;
    }
}

static BOOL is_checked(HWND checkbox) {
    return SendMessageW(checkbox, BM_GETCHECK, 0, 0) == BST_CHECKED;
}

static void set_checked(HWND checkbox, BOOL checked) {
    SendMessageW(checkbox, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
}

static BOOL is_our_window(const AppState *app, HWND hwnd) {
    if (!app || !hwnd) {
        return FALSE;
    }

    if (hwnd == app->main_hwnd || hwnd == app->float_hwnd) {
        return TRUE;
    }

    if (app->main_hwnd && IsChild(app->main_hwnd, hwnd)) {
        return TRUE;
    }

    if (app->float_hwnd && IsChild(app->float_hwnd, hwnd)) {
        return TRUE;
    }

    return FALSE;
}

static void set_status(AppState *app, const wchar_t *text) {
    if (!app || !text) {
        return;
    }

    if (app->status_label) {
        SetWindowTextW(app->status_label, text);
    }

    if (app->float_status) {
        SetWindowTextW(app->float_status, text);
    }
}

static void append_hotkey_token(wchar_t *buffer, size_t buffer_len, const wchar_t *token) {
    if (buffer[0] != L'\0') {
        wcscat_s(buffer, buffer_len, L"+");
    }
    wcscat_s(buffer, buffer_len, token);
}

static void hotkey_vk_to_text(UINT vk, wchar_t *out, size_t out_len) {
    if (vk >= 'A' && vk <= 'Z') {
        out[0] = (wchar_t)vk;
        out[1] = L'\0';
        return;
    }

    if (vk >= '0' && vk <= '9') {
        out[0] = (wchar_t)vk;
        out[1] = L'\0';
        return;
    }

    if (vk >= VK_F1 && vk <= VK_F24) {
        swprintf(out, out_len, L"F%u", (unsigned)(vk - VK_F1 + 1));
        return;
    }

    switch (vk) {
    case VK_SPACE:
        wcscpy_s(out, out_len, L"Space");
        return;
    case VK_RETURN:
        wcscpy_s(out, out_len, L"Enter");
        return;
    case VK_TAB:
        wcscpy_s(out, out_len, L"Tab");
        return;
    case VK_ESCAPE:
        wcscpy_s(out, out_len, L"Esc");
        return;
    default:
        swprintf(out, out_len, L"VK_%u", (unsigned)vk);
        return;
    }
}

static void format_hotkey_text(UINT mods, UINT vk, wchar_t *out, size_t out_len) {
    wchar_t key_text[32];

    out[0] = L'\0';

    if (mods & MOD_CONTROL) {
        append_hotkey_token(out, out_len, L"Ctrl");
    }
    if (mods & MOD_ALT) {
        append_hotkey_token(out, out_len, L"Alt");
    }
    if (mods & MOD_SHIFT) {
        append_hotkey_token(out, out_len, L"Shift");
    }
    if (mods & MOD_WIN) {
        append_hotkey_token(out, out_len, L"Win");
    }

    hotkey_vk_to_text(vk, key_text, _countof(key_text));
    append_hotkey_token(out, out_len, key_text);
}

static UINT parse_hotkey_key(const wchar_t *input) {
    wchar_t key[32];
    size_t len = 0;
    size_t i = 0;

    if (!input) {
        return 0;
    }

    wcsncpy_s(key, _countof(key), input, _TRUNCATE);
    trim_wide_whitespace(key);
    len = wcslen(key);

    if (len == 0) {
        return 0;
    }

    for (i = 0; i < len; ++i) {
        key[i] = towupper(key[i]);
    }

    if (len == 1) {
        SHORT mapped = VkKeyScanW(key[0]);
        if (mapped == -1) {
            return 0;
        }
        return (UINT)LOBYTE(mapped);
    }

    if (key[0] == L'F' && len <= 3) {
        int fn = _wtoi(key + 1);
        if (fn >= 1 && fn <= 24) {
            return VK_F1 + (UINT)(fn - 1);
        }
    }

    if (wcscmp(key, L"SPACE") == 0) {
        return VK_SPACE;
    }
    if (wcscmp(key, L"ENTER") == 0) {
        return VK_RETURN;
    }
    if (wcscmp(key, L"TAB") == 0) {
        return VK_TAB;
    }
    if (wcscmp(key, L"ESC") == 0 || wcscmp(key, L"ESCAPE") == 0) {
        return VK_ESCAPE;
    }

    return 0;
}

static UINT read_modifiers(const AppState *app) {
    UINT mods = 0;

    if (is_checked(app->check_ctrl)) {
        mods |= MOD_CONTROL;
    }
    if (is_checked(app->check_alt)) {
        mods |= MOD_ALT;
    }
    if (is_checked(app->check_shift)) {
        mods |= MOD_SHIFT;
    }
    if (is_checked(app->check_win)) {
        mods |= MOD_WIN;
    }

    return mods;
}

static void update_hotkey_preview(AppState *app) {
    wchar_t hotkey_text[64];
    wchar_t line[96];

    if (!app || !app->current_hotkey_label) {
        return;
    }

    format_hotkey_text(app->hotkey_mods, app->hotkey_vk, hotkey_text, _countof(hotkey_text));
    swprintf(line, _countof(line), L"当前快捷键：%ls", hotkey_text);
    SetWindowTextW(app->current_hotkey_label, line);
}

static void update_float_button(AppState *app) {
    if (!app || !app->float_hwnd) {
        return;
    }
    InvalidateRect(app->float_hwnd, NULL, TRUE);
}

static DWORD clamp_setting(DWORD value, DWORD min_value, DWORD max_value) {
    if (value < min_value) {
        return min_value;
    }
    if (value > max_value) {
        return max_value;
    }
    return value;
}

static DWORD read_edit_dword(HWND edit, DWORD fallback, DWORD min_value, DWORD max_value) {
    wchar_t text[64];
    wchar_t *end_ptr = NULL;
    unsigned long parsed = 0;

    if (!edit) {
        return fallback;
    }

    GetWindowTextW(edit, text, _countof(text));
    trim_wide_whitespace(text);
    if (text[0] == L'\0') {
        return clamp_setting(fallback, min_value, max_value);
    }

    parsed = wcstoul(text, &end_ptr, 10);
    if (end_ptr == text) {
        return clamp_setting(fallback, min_value, max_value);
    }

    return clamp_setting((DWORD)parsed, min_value, max_value);
}

static void write_dword_to_edit(HWND edit, DWORD value) {
    wchar_t text[64];

    if (!edit) {
        return;
    }

    swprintf(text, _countof(text), L"%lu", (unsigned long)value);
    SetWindowTextW(edit, text);
}

static void set_default_recorder_config(AppState *app) {
    if (!app) {
        return;
    }

    app->recorder_config.sample_rate = 16000;
    app->recorder_config.channels = 1;
    app->recorder_config.bits_per_sample = 16;
    app->recorder_config.voice_threshold = 1400;
    app->recorder_config.silence_timeout_ms = 1500;
    app->recorder_config.min_record_ms = 900;
    app->recorder_config.max_record_ms = 30000;
}

static void try_auto_fill_sherpa_defaults(AppState *app) {
    wchar_t exe_path[MAX_PATH];
    wchar_t app_dir[MAX_PATH];
    wchar_t parent_dir[MAX_PATH];
    wchar_t grandparent_dir[MAX_PATH];
    const wchar_t *roots[3];
    int i;

    if (!app || app->backend != ASR_BACKEND_SHERPA) {
        return;
    }

    if (app->sherpa_exe[0] != L'\0' && app->sherpa_args[0] != L'\0') {
        return;
    }

    app_dir[0] = L'\0';
    parent_dir[0] = L'\0';
    grandparent_dir[0] = L'\0';

    GetModuleFileNameW(NULL, exe_path, MAX_PATH);
    extract_parent_dir(exe_path, app_dir, _countof(app_dir));
    if (app_dir[0] != L'\0') {
        extract_parent_dir(app_dir, parent_dir, _countof(parent_dir));
    }
    if (parent_dir[0] != L'\0') {
        extract_parent_dir(parent_dir, grandparent_dir, _countof(grandparent_dir));
    }

    roots[0] = app_dir;
    roots[1] = parent_dir;
    roots[2] = grandparent_dir;

    for (i = 0; i < 3; ++i) {
        wchar_t exe_path[MAX_PATH];
        wchar_t exe_path_cuda[MAX_PATH];
        wchar_t tokens_path[MAX_PATH];
        wchar_t model_path[MAX_PATH];
        BOOL is_cuda = FALSE;

        if (!roots[i] || roots[i][0] == L'\0') {
            continue;
        }

        swprintf(exe_path_cuda, _countof(exe_path_cuda), L"%ls\\third_party\\sherpa\\sherpa-onnx-v1.12.29-win-x64-cuda\\bin\\sherpa-onnx-offline.exe", roots[i]);
        swprintf(exe_path, _countof(exe_path), L"%ls\\third_party\\sherpa\\sherpa-onnx-v1.12.29-win-x64-static-MT-Release-no-tts\\bin\\sherpa-onnx-offline.exe", roots[i]);
        
        swprintf(tokens_path, _countof(tokens_path), L"%ls\\third_party\\sherpa\\models\\paraformer-zh\\tokens.txt", roots[i]);
        swprintf(model_path, _countof(model_path), L"%ls\\third_party\\sherpa\\models\\paraformer-zh\\model.int8.onnx", roots[i]);

        if (file_exists_non_dir(exe_path_cuda)) {
            wcsncpy_s(exe_path, _countof(exe_path), exe_path_cuda, _TRUNCATE);
            is_cuda = TRUE;
        } else if (!file_exists_non_dir(exe_path)) {
            continue;
        }

        if (!file_exists_non_dir(tokens_path) || !file_exists_non_dir(model_path)) {
            continue;
        }

        wcsncpy_s(app->sherpa_exe, _countof(app->sherpa_exe), exe_path, _TRUNCATE);
        if (is_cuda) {
            swprintf(app->sherpa_args, _countof(app->sherpa_args), L"--provider=cuda --paraformer=%ls --tokens=%ls --num-threads=2 --decoding-method=greedy_search", model_path, tokens_path);
        } else {
            swprintf(app->sherpa_args, _countof(app->sherpa_args), L"--paraformer=%ls --tokens=%ls --num-threads=2 --decoding-method=greedy_search", model_path, tokens_path);
        }

        app_log_line(app, "auto-filled sherpa defaults from local third_party folder");
        return;
    }
}

static void sync_runtime_settings_to_ui(AppState *app) {
    if (!app) {
        return;
    }

    if (app->backend_combo) {
        if (app->backend == ASR_BACKEND_SHERPA) {
            SendMessageW(app->backend_combo, CB_SETCURSEL, 1, 0);
        } else if (app->backend == ASR_BACKEND_GLADIA) {
            SendMessageW(app->backend_combo, CB_SETCURSEL, 2, 0);
        } else if (app->backend == ASR_BACKEND_GEMINI) {
            SendMessageW(app->backend_combo, CB_SETCURSEL, 3, 0);
        } else {
            SendMessageW(app->backend_combo, CB_SETCURSEL, 0, 0);
        }
    }

    if (app->lang_combo) {
        LRESULT idx = SendMessageW(app->lang_combo, CB_FINDSTRINGEXACT, -1, (LPARAM)app->target_lang);
        if (idx != CB_ERR) {
            SendMessageW(app->lang_combo, CB_SETCURSEL, idx, 0);
        } else {
            SendMessageW(app->lang_combo, CB_SETCURSEL, 0, 0);
        }
    }

    if (app->gemini_model_combo) {
        LRESULT idx = SendMessageW(app->gemini_model_combo, CB_FINDSTRINGEXACT, -1, (LPARAM)app->gemini_model);
        if (idx != CB_ERR) {
            SendMessageW(app->gemini_model_combo, CB_SETCURSEL, idx, 0);
        } else {
            SendMessageW(app->gemini_model_combo, CB_SETCURSEL, 0, 0);
        }
    }

    if (app->thinking_combo) {
        LRESULT idx = SendMessageW(app->thinking_combo, CB_FINDSTRINGEXACT, -1, (LPARAM)app->thinking_level);
        if (idx != CB_ERR) {
            SendMessageW(app->thinking_combo, CB_SETCURSEL, idx, 0);
        } else {
            SendMessageW(app->thinking_combo, CB_SETCURSEL, 1, 0); // Default to LOW
        }
    }

    if (app->continuous_check) {
        set_checked(app->continuous_check, app->continuous_mode);
    }

    if (app->auto_stop_check) {
        set_checked(app->auto_stop_check, app->auto_stop_enabled);
    }

    if (app->translate_check) {
        set_checked(app->translate_check, app->translate_enabled);
    }

    if (app->model_combo) {
        SendMessageW(app->model_combo, CB_SETCURSEL, app->local_model_index, 0);
    }

    if (app->sherpa_exe_edit) {
        SetWindowTextW(app->sherpa_exe_edit, app->sherpa_exe);
    }

    if (app->sherpa_args_edit) {
        SetWindowTextW(app->sherpa_args_edit, app->sherpa_args);
    }

    if (app->gemini_prompt_edit) {
        SetWindowTextW(app->gemini_prompt_edit, app->gemini_prompt);
    }

    if (app->replace_rules_edit) {
        SetWindowTextW(app->replace_rules_edit, app->replace_rules);
    }

    write_dword_to_edit(app->threshold_edit, (DWORD)app->recorder_config.voice_threshold);
    write_dword_to_edit(app->silence_edit, app->recorder_config.silence_timeout_ms);
    write_dword_to_edit(app->minrec_edit, app->recorder_config.min_record_ms);
    write_dword_to_edit(app->maxrec_edit, app->recorder_config.max_record_ms);
}

static BOOL apply_runtime_settings_from_ui(AppState *app, BOOL persist) {
    LRESULT selected_backend = 0;
    LRESULT selected_lang = 0;
    LRESULT selected_model = 0;

    if (!app) {
        return FALSE;
    }

    if (app->backend_combo) {
        selected_backend = SendMessageW(app->backend_combo, CB_GETCURSEL, 0, 0);
        if (selected_backend == 1) {
            app->backend = ASR_BACKEND_SHERPA;
        } else if (selected_backend == 2) {
            app->backend = ASR_BACKEND_GLADIA;
        } else if (selected_backend == 3) {
            app->backend = ASR_BACKEND_GEMINI;
        } else {
            app->backend = ASR_BACKEND_GROQ;
        }
    }

    app->continuous_mode = app->continuous_check ? is_checked(app->continuous_check) : FALSE;
    app->auto_stop_enabled = app->auto_stop_check ? is_checked(app->auto_stop_check) : TRUE;
    app->translate_enabled = app->translate_check ? is_checked(app->translate_check) : FALSE;

    if (app->model_combo) {
        LRESULT sel = SendMessageW(app->model_combo, CB_GETCURSEL, 0, 0);
        if (sel != CB_ERR) {
            app->local_model_index = (int)sel;
        }
    }

    if (app->lang_combo) {
        selected_lang = SendMessageW(app->lang_combo, CB_GETCURSEL, 0, 0);
        if (selected_lang != CB_ERR) {
            SendMessageW(app->lang_combo, CB_GETLBTEXT, (WPARAM)selected_lang, (LPARAM)app->target_lang);
        }
    }

    if (app->gemini_model_combo) {
        selected_model = SendMessageW(app->gemini_model_combo, CB_GETCURSEL, 0, 0);
        if (selected_model != CB_ERR) {
            SendMessageW(app->gemini_model_combo, CB_GETLBTEXT, (WPARAM)selected_model, (LPARAM)app->gemini_model);
        }
    }

    if (app->thinking_combo) {
        LRESULT selected_thinking = SendMessageW(app->thinking_combo, CB_GETCURSEL, 0, 0);
        if (selected_thinking != CB_ERR) {
            SendMessageW(app->thinking_combo, CB_GETLBTEXT, (WPARAM)selected_thinking, (LPARAM)app->thinking_level);
        }
    }

    if (app->gemini_key_edit) {
        GetWindowTextW(app->gemini_key_edit, app->gemini_key, _countof(app->gemini_key));
        trim_wide_whitespace(app->gemini_key);
    }
    
    if (app->project_id_edit) {
        GetWindowTextW(app->project_id_edit, app->project_id, _countof(app->project_id));
        trim_wide_whitespace(app->project_id);
    }
    
    if (app->gladia_key_edit) {
        GetWindowTextW(app->gladia_key_edit, app->gladia_key, _countof(app->gladia_key));
        trim_wide_whitespace(app->gladia_key);
    }

    if (app->gemini_prompt_edit) {
        GetWindowTextW(app->gemini_prompt_edit, app->gemini_prompt, _countof(app->gemini_prompt));
        trim_wide_whitespace(app->gemini_prompt);
    }

    if (app->sherpa_exe_edit) {
        GetWindowTextW(app->sherpa_exe_edit, app->sherpa_exe, _countof(app->sherpa_exe));
        trim_wide_whitespace(app->sherpa_exe);
    }

    if (app->sherpa_args_edit) {
        GetWindowTextW(app->sherpa_args_edit, app->sherpa_args, _countof(app->sherpa_args));
        trim_wide_whitespace(app->sherpa_args);
    }

    if (app->replace_rules_edit) {
        GetWindowTextW(app->replace_rules_edit, app->replace_rules, _countof(app->replace_rules));
        trim_wide_whitespace(app->replace_rules);
    }

    app->recorder_config.sample_rate = clamp_setting(app->recorder_config.sample_rate, 8000, 48000);
    app->recorder_config.channels = 1;
    app->recorder_config.bits_per_sample = 16;
    app->recorder_config.voice_threshold = (SHORT)read_edit_dword(
        app->threshold_edit,
        (DWORD)app->recorder_config.voice_threshold,
        120,
        6000);
    app->recorder_config.silence_timeout_ms = read_edit_dword(
        app->silence_edit,
        app->recorder_config.silence_timeout_ms,
        400,
        6000);
    app->recorder_config.min_record_ms = read_edit_dword(
        app->minrec_edit,
        app->recorder_config.min_record_ms,
        300,
        5000);
    app->recorder_config.max_record_ms = read_edit_dword(
        app->maxrec_edit,
        app->recorder_config.max_record_ms,
        3000,
        120000);

    sync_runtime_settings_to_ui(app);

    if (persist) {
        save_settings(app);
    }

    return TRUE;
}

static void set_selfcheck_text(AppState *app, const wchar_t *text) {
    if (app && app->selfcheck_label && text) {
        SetWindowTextW(app->selfcheck_label, text);
    }
}

static void append_report_line(wchar_t *report, size_t report_len, const wchar_t *line) {
    if (!report || report_len == 0 || !line) {
        return;
    }

    if (report[0] != L'\0') {
        wcscat_s(report, report_len, L"\r\n");
    }
    wcscat_s(report, report_len, line);
}

static int validate_sherpa_path_option(const wchar_t *args,
                                       const wchar_t *prefix,
                                       const wchar_t *base_dir,
                                       wchar_t *report,
                                       size_t report_len) {
    wchar_t value[1024];
    wchar_t resolved[2048];
    wchar_t line[1200];

    if (!extract_option_value(args, prefix, value, _countof(value))) {
        swprintf(line, _countof(line), L"缺少参数：%ls<路径>", prefix);
        append_report_line(report, report_len, line);
        return 1;
    }

    resolve_path_from_base(base_dir, value, resolved, _countof(resolved));
    if (!file_exists_non_dir(resolved)) {
        swprintf(line, _countof(line), L"文件不存在：%ls -> %ls", prefix, resolved);
        append_report_line(report, report_len, line);
        return 1;
    }

    return 0;
}

static BOOL build_self_check_report(AppState *app, wchar_t *report, size_t report_len) {
    int issues = 0;
    wchar_t hotkey_text[32];

    if (!app || !report || report_len == 0) {
        return FALSE;
    }

    report[0] = L'\0';
    apply_runtime_settings_from_ui(app, FALSE);

    GetWindowTextW(app->hotkey_edit, hotkey_text, _countof(hotkey_text));
    if (parse_hotkey_key(hotkey_text) == 0) {
        append_report_line(report, report_len, L"快捷键主键无效。");
        issues++;
    }
    if (read_modifiers(app) == 0) {
        append_report_line(report, report_len, L"至少勾选一个快捷键修饰键（Ctrl/Alt/Shift/Win）。");
        issues++;
    }

    if (app->backend == ASR_BACKEND_GROQ) {
        wchar_t api_key[512];
        GetWindowTextW(app->api_edit, api_key, _countof(api_key));
        trim_wide_whitespace(api_key);
        if (api_key[0] == L'\0') {
            append_report_line(report, report_len, L"Groq API Key 为空。");
            issues++;
        }
    } else {
        wchar_t sherpa_dir[MAX_PATH];
        BOOL has_paraformer = FALSE;
        BOOL has_whisper_encoder = FALSE;
        BOOL has_whisper_decoder = FALSE;
        wchar_t tmp[64];

        if (app->sherpa_exe[0] == L'\0') {
            append_report_line(report, report_len, L"Sherpa 可执行程序路径为空。");
            issues++;
        } else if (!file_exists_non_dir(app->sherpa_exe)) {
            append_report_line(report, report_len, L"Sherpa 可执行程序不存在。");
            issues++;
        }

        if (app->sherpa_args[0] == L'\0') {
            append_report_line(report, report_len, L"Sherpa 参数为空。");
            issues++;
        }

        extract_parent_dir(app->sherpa_exe, sherpa_dir, _countof(sherpa_dir));
        has_paraformer = extract_option_value(app->sherpa_args, L"--paraformer=", tmp, _countof(tmp));
        BOOL has_funasr = extract_option_value(app->sherpa_args, L"--funasr-nano-llm=", tmp, _countof(tmp));
        BOOL has_sense_voice = extract_option_value(app->sherpa_args, L"--sense-voice-model=", tmp, _countof(tmp));
        has_whisper_encoder = extract_option_value(app->sherpa_args, L"--whisper-encoder=", tmp, _countof(tmp));
        has_whisper_decoder = extract_option_value(app->sherpa_args, L"--whisper-decoder=", tmp, _countof(tmp));

        if (app->sherpa_args[0] != L'\0') {
            if (!has_funasr) {
                issues += validate_sherpa_path_option(app->sherpa_args, L"--tokens=", sherpa_dir, report, report_len);
            }

            if (has_paraformer) {
                issues += validate_sherpa_path_option(app->sherpa_args, L"--paraformer=", sherpa_dir, report, report_len);
            } else if (has_funasr) {
                issues += validate_sherpa_path_option(app->sherpa_args, L"--funasr-nano-encoder-adaptor=", sherpa_dir, report, report_len);
                issues += validate_sherpa_path_option(app->sherpa_args, L"--funasr-nano-llm=", sherpa_dir, report, report_len);
                issues += validate_sherpa_path_option(app->sherpa_args, L"--funasr-nano-embedding=", sherpa_dir, report, report_len);
            } else if (has_sense_voice) {
                issues += validate_sherpa_path_option(app->sherpa_args, L"--sense-voice-model=", sherpa_dir, report, report_len);
            } else if (has_whisper_encoder || has_whisper_decoder) {
                issues += validate_sherpa_path_option(app->sherpa_args, L"--whisper-encoder=", sherpa_dir, report, report_len);
                issues += validate_sherpa_path_option(app->sherpa_args, L"--whisper-decoder=", sherpa_dir, report, report_len);
            } else {
                issues += validate_sherpa_path_option(app->sherpa_args, L"--encoder=", sherpa_dir, report, report_len);
                issues += validate_sherpa_path_option(app->sherpa_args, L"--decoder=", sherpa_dir, report, report_len);
                issues += validate_sherpa_path_option(app->sherpa_args, L"--joiner=", sherpa_dir, report, report_len);
            }
        }
    }

    if (issues == 0) {
        wcsncpy_s(report, report_len, L"自检通过：当前配置可用。", _TRUNCATE);
        return TRUE;
    }

    return FALSE;
}

static void run_self_check(AppState *app, BOOL popup) {
    wchar_t report[4096];
    BOOL ok = build_self_check_report(app, report, _countof(report));

    if (!app) {
        return;
    }

    set_selfcheck_text(app, report);
    if (ok) {
        set_status(app, L"自检通过。");
        app_log_line(app, "self-check ok");
    } else {
        set_status(app, L"自检发现问题。");
        app_log_line(app, "self-check failed");
    }

    if (popup) {
        MessageBoxW(app->main_hwnd,
                    report,
                    ok ? L"自检通过" : L"自检问题",
                    MB_OK | (ok ? MB_ICONINFORMATION : MB_ICONWARNING));
    }
}

static void apply_model_selection(AppState *app, int sel) {
    wchar_t exe_path[MAX_PATH];
    wchar_t app_dir[MAX_PATH];
    wchar_t parent_dir[MAX_PATH];
    wchar_t grandparent_dir[MAX_PATH];
    wchar_t script_path[2048];
    wchar_t params[4096];
    const wchar_t *roots[3];
    int i;
    const wchar_t *work_dir = NULL;
    HINSTANCE shell_result;
    
    GetModuleFileNameW(NULL, exe_path, MAX_PATH);
    extract_parent_dir(exe_path, app_dir, _countof(app_dir));
    
    parent_dir[0] = L'\0';
    grandparent_dir[0] = L'\0';
    if (app_dir[0] != L'\0') {
        extract_parent_dir(app_dir, parent_dir, _countof(parent_dir));
    }
    if (parent_dir[0] != L'\0') {
        extract_parent_dir(parent_dir, grandparent_dir, _countof(grandparent_dir));
    }

    roots[0] = app_dir;
    roots[1] = parent_dir;
    roots[2] = grandparent_dir;

    // Find the correct root that contains scripts\install_model.bat
    wchar_t correct_root[MAX_PATH] = {0};
    for (i = 0; i < 3; ++i) {
        if (!roots[i] || roots[i][0] == L'\0') continue;
        swprintf(script_path, _countof(script_path), L"%ls\\scripts\\install_model.bat", roots[i]);
        if (file_exists_non_dir(script_path)) {
            wcsncpy_s(correct_root, _countof(correct_root), roots[i], _TRUNCATE);
            break;
        }
    }

    if (correct_root[0] == L'\0') {
        // If not found, fallback to the grandparent directory (assuming we run in build/Release)
        wcsncpy_s(correct_root, _countof(correct_root), grandparent_dir, _TRUNCATE);
        swprintf(script_path, _countof(script_path), L"%ls\\scripts\\install_model.bat", correct_root);
    }

    const wchar_t* model_id = L"paraformer";
    if (sel == 1) model_id = L"zipformer";
    if (sel == 2) model_id = L"funasr";

    // Fill default arguments according to the selection using the correct root
    wchar_t sherpa_exe[2048];
    BOOL is_cuda = FALSE;
    swprintf(sherpa_exe, _countof(sherpa_exe), L"%ls\\third_party\\sherpa\\sherpa-onnx-v1.12.29-win-x64-cuda\\bin\\sherpa-onnx-offline.exe", correct_root);
    if (file_exists_non_dir(sherpa_exe)) {
        is_cuda = TRUE;
    } else {
        swprintf(sherpa_exe, _countof(sherpa_exe), L"%ls\\third_party\\sherpa\\sherpa-onnx-v1.12.29-win-x64-static-MT-Release-no-tts\\bin\\sherpa-onnx-offline.exe", correct_root);
    }
    
    wcsncpy_s(app->sherpa_exe, _countof(app->sherpa_exe), sherpa_exe, _TRUNCATE);
    
    wchar_t cuda_prefix[32] = L"";
    if (is_cuda) {
        wcscpy_s(cuda_prefix, _countof(cuda_prefix), L"--provider=cuda ");
    }
    
    if (sel == 0) {
        swprintf(app->sherpa_args, _countof(app->sherpa_args), L"%ls--paraformer=%ls\\third_party\\sherpa\\models\\paraformer-zh\\model.int8.onnx --tokens=%ls\\third_party\\sherpa\\models\\paraformer-zh\\tokens.txt --num-threads=2 --decoding-method=greedy_search", cuda_prefix, correct_root, correct_root);
    } else if (sel == 1) {
        swprintf(app->sherpa_args, _countof(app->sherpa_args), L"%ls--encoder=\"%ls\\third_party\\sherpa\\models\\zipformer-zh\\encoder.int8.onnx\" --decoder=\"%ls\\third_party\\sherpa\\models\\zipformer-zh\\decoder.onnx\" --joiner=\"%ls\\third_party\\sherpa\\models\\zipformer-zh\\joiner.int8.onnx\" --tokens=\"%ls\\third_party\\sherpa\\models\\zipformer-zh\\tokens.txt\" --num-threads=2 --decoding-method=greedy_search", cuda_prefix, correct_root, correct_root, correct_root, correct_root);
    } else if (sel == 2) {
        swprintf(app->sherpa_args, _countof(app->sherpa_args), L"%ls--funasr-nano-encoder-adaptor=\"%ls\\third_party\\sherpa\\models\\funasr\\encoder_adaptor.int8.onnx\" --funasr-nano-llm=\"%ls\\third_party\\sherpa\\models\\funasr\\llm.int8.onnx\" --funasr-nano-embedding=\"%ls\\third_party\\sherpa\\models\\funasr\\embedding.int8.onnx\" --funasr-nano-tokenizer=\"%ls\\third_party\\sherpa\\models\\funasr\\Qwen3-0.6B\" --tokens=\"%ls\\third_party\\sherpa\\models\\funasr\\tokens.txt\"", cuda_prefix, correct_root, correct_root, correct_root, correct_root, correct_root);
    }
    

    // Automatically setup hotwords
    {
        wchar_t hotwords_path[MAX_PATH];
        wchar_t config_dir[MAX_PATH];
        extract_parent_dir(app->config_path, config_dir, _countof(config_dir));
        swprintf(hotwords_path, _countof(hotwords_path), L"%ls\\hotwords.txt", config_dir);

        if (!file_exists_non_dir(hotwords_path)) {
            FILE *f = _wfopen(hotwords_path, L"wt,ccs=UTF-8");
            if (f) {
                fputws(L"你好 1.5\n", f);
                fputws(L"语音输入 2.0\n", f);
                fclose(f);
            }
        }

        if (sel == 1) { // Zipformer
            wchar_t append_args[1024];
            wchar_t bpe_path[MAX_PATH];
            swprintf(bpe_path, _countof(bpe_path), L"%ls\\third_party\\sherpa\\models\\zipformer-zh\\bpe.model", correct_root);
            swprintf(append_args, _countof(append_args), L" --hotwords-file=\"%ls\" --hotwords-score=1.5 --bpe-model=\"%ls\" --modeling-unit=cjkchar+bpe", hotwords_path, bpe_path);
            wcscat_s(app->sherpa_args, _countof(app->sherpa_args), append_args);
        } else if (sel == 2) { // FunASR
            wchar_t append_args[1024];
            FILE *f = _wfopen(hotwords_path, L"rt,ccs=UTF-8");
            if (f) {
                wchar_t line[256];
                wchar_t all_words[1024] = L"";
                int first = 1;
                while (fgetws(line, _countof(line), f)) {
                    size_t len = wcslen(line);
                    while(len > 0 && (line[len-1] == L'\n' || line[len-1] == L'\r' || line[len-1] == L' ')) {
                        line[len-1] = L'\0';
                        len--;
                    }
                    if (len > 0 && wcschr(line, L'.') == NULL) {
                        if (!first) wcscat_s(all_words, _countof(all_words), L",");
                        wcscat_s(all_words, _countof(all_words), line);
                        first = 0;
                    }
                }
                fclose(f);
                if (all_words[0] != L'\0') {
                    swprintf(append_args, _countof(append_args), L" --funasr-nano-hotwords=\"%ls\"", all_words);
                    wcscat_s(app->sherpa_args, _countof(app->sherpa_args), append_args);
                }
            }
        }
    }

    SetWindowTextW(app->sherpa_exe_edit, app->sherpa_exe);
    SetWindowTextW(app->sherpa_args_edit, app->sherpa_args);
    save_settings(app);

    swprintf(params, _countof(params), L"/c \"%ls\" %ls", script_path, model_id);
    shell_result = ShellExecuteW(app->main_hwnd, L"open", L"cmd.exe", params, NULL, SW_SHOWNORMAL);
    
    if ((INT_PTR)shell_result <= 32) {
        set_status(app, L"已更新配置，但启动下载脚本失败。");
    } else {
        set_status(app, L"正在下载模型，并已更新配置。");
    }
}

static BOOL launch_sherpa_installer(AppState *app) {
    wchar_t exe_path[MAX_PATH];
    wchar_t app_dir[MAX_PATH];
    wchar_t parent_dir[MAX_PATH];
    wchar_t grandparent_dir[MAX_PATH];
    wchar_t script_path[2048];
    wchar_t candidate[2048];
    wchar_t params[4096];
    const wchar_t *work_dir = NULL;
    HINSTANCE shell_result;

    if (!app) {
        return FALSE;
    }

    apply_runtime_settings_from_ui(app, FALSE);
    app->backend = ASR_BACKEND_SHERPA;
    try_auto_fill_sherpa_defaults(app);
    sync_runtime_settings_to_ui(app);

    if (app->sherpa_exe[0] != L'\0' && app->sherpa_args[0] != L'\0') {
        save_settings(app);
        run_self_check(app, FALSE);
        set_status(app, L"已自动应用本地 Sherpa 配置。");
        return TRUE;
    }

    GetModuleFileNameW(NULL, exe_path, MAX_PATH);
    extract_parent_dir(exe_path, app_dir, _countof(app_dir));
    if (app_dir[0] == L'\0') {
        set_status(app, L"无法定位程序目录。");
        return FALSE;
    }

    parent_dir[0] = L'\0';
    grandparent_dir[0] = L'\0';
    extract_parent_dir(app_dir, parent_dir, _countof(parent_dir));
    if (parent_dir[0] != L'\0') {
        extract_parent_dir(parent_dir, grandparent_dir, _countof(grandparent_dir));
    }

    swprintf(script_path, _countof(script_path), L"%ls\\scripts\\install_sherpa.bat", app_dir);
    if (!file_exists_non_dir(script_path)) {
        if (parent_dir[0] != L'\0') {
            swprintf(candidate, _countof(candidate), L"%ls\\scripts\\install_sherpa.bat", parent_dir);
            if (file_exists_non_dir(candidate)) {
                wcsncpy_s(script_path, _countof(script_path), candidate, _TRUNCATE);
            }
        }
    }

    if (!file_exists_non_dir(script_path)) {
        if (grandparent_dir[0] != L'\0') {
            swprintf(candidate, _countof(candidate), L"%ls\\scripts\\install_sherpa.bat", grandparent_dir);
            if (file_exists_non_dir(candidate)) {
                wcsncpy_s(script_path, _countof(script_path), candidate, _TRUNCATE);
            }
        }
    }

    if (!file_exists_non_dir(script_path)) {
        set_status(app, L"未找到安装脚本 scripts\\install_sherpa.bat。");
        return FALSE;
    }

    if (wcsncmp(script_path, app_dir, wcslen(app_dir)) == 0) {
        work_dir = app_dir;
    } else if (parent_dir[0] != L'\0' && wcsncmp(script_path, parent_dir, wcslen(parent_dir)) == 0) {
        work_dir = parent_dir;
    } else if (grandparent_dir[0] != L'\0' && wcsncmp(script_path, grandparent_dir, wcslen(grandparent_dir)) == 0) {
        work_dir = grandparent_dir;
    } else {
        work_dir = app_dir;
    }

    swprintf(params, _countof(params), L"/c \"%ls\"", script_path);

    shell_result = ShellExecuteW(app->main_hwnd,
                                 L"open",
                                 L"cmd.exe",
                                 params,
                                 work_dir,
                                 SW_SHOWNORMAL);
    if ((INT_PTR)shell_result <= 32) {
        set_status(app, L"启动 Sherpa 安装脚本失败。");
        return FALSE;
    }

    set_status(app, L"已在新窗口启动 Sherpa 安装脚本。");
    app_log_line(app, "sherpa installer launched");
    return TRUE;
}

static void save_settings(AppState *app) {
    wchar_t api_key[512];
    wchar_t key_text[32];
    wchar_t replace_rules[2048];
    wchar_t content[16384];
    UINT mods = 0;
    int bytes_needed = 0;
    char *mb_content = NULL;
    HANDLE file_handle = INVALID_HANDLE_VALUE;
    DWORD written = 0;

    if (!app) {
        return;
    }

    replace_rules[0] = L'\0';
    GetWindowTextW(app->api_edit, api_key, _countof(api_key));
    GetWindowTextW(app->hotkey_edit, key_text, _countof(key_text));
    if (app->replace_rules_edit) {
        GetWindowTextW(app->replace_rules_edit, replace_rules, _countof(replace_rules));
    } else {
        wcsncpy_s(replace_rules, _countof(replace_rules), app->replace_rules, _TRUNCATE);
    }
    wcsncpy_s(app->replace_rules, _countof(app->replace_rules), replace_rules, _TRUNCATE);
    trim_wide_whitespace(app->replace_rules);

    mods = read_modifiers(app);
    if (mods == 0) {
        mods = app->hotkey_mods ? app->hotkey_mods : (MOD_CONTROL | MOD_ALT);
    }

    swprintf(content,
             _countof(content),
             L"[settings]\r\n"
             L"api_key=%ls\r\n"
             L"hotkey_key=%ls\r\n"
             L"hotkey_mods=%u\r\n"
             L"backend=%ls\r\n"
             L"sherpa_exe=%ls\r\n"
             L"sherpa_args=%ls\r\n"
             L"replace_rules=%ls\r\n"
             L"continuous_mode=%u\r\n"
             L"auto_stop=%u\r\n"
             L"sample_rate=%u\r\n"
             L"voice_threshold=%d\r\n"
             L"silence_timeout_ms=%u\r\n"
             L"min_record_ms=%u\r\n"
             L"max_record_ms=%u\r\n"
             L"mic_device_id=%u\r\n"
             L"gemini_key=%ls\r\n"
             L"project_id=%ls\r\n"
             L"gemini_model=%ls\r\n"
             L"gemini_prompt=%ls\r\n"
             L"thinking_level=%ls\r\n"
             L"gladia_key=%ls\r\n"
             L"target_lang=%ls\r\n"
             L"translate_enabled=%u\r\n",
             api_key,
             key_text,
             (unsigned)mods,
             asr_backend_name(app->backend),
             app->sherpa_exe,
             app->sherpa_args,
             app->replace_rules,
             app->continuous_mode ? 1u : 0u,
             app->auto_stop_enabled ? 1u : 0u,
             (unsigned)app->recorder_config.sample_rate,
             (int)app->recorder_config.voice_threshold,
             (unsigned)app->recorder_config.silence_timeout_ms,
             (unsigned)app->recorder_config.min_record_ms,
             (unsigned)app->recorder_config.max_record_ms,
             app->mic_device_id,
             app->gemini_key,
             app->project_id,
             app->gemini_model,
             app->gemini_prompt,
             app->thinking_level,
             app->gladia_key,
             app->target_lang,
             app->translate_enabled ? 1u : 0u,
             app->local_model_index);

    file_handle = CreateFileW(app->config_path,
                              GENERIC_WRITE,
                              FILE_SHARE_READ,
                              NULL,
                              CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL,
                              NULL);
    if (file_handle == INVALID_HANDLE_VALUE) {
        return;
    }

    // Write UTF-16 LE BOM
    unsigned char bom[] = { 0xFF, 0xFE };
    WriteFile(file_handle, bom, sizeof(bom), &written, NULL);

    // Write wide character content directly
    DWORD bytes_to_write = (DWORD)(wcslen(content) * sizeof(wchar_t));
    WriteFile(file_handle, content, bytes_to_write, &written, NULL);
    CloseHandle(file_handle);
}

static void load_settings(AppState *app) {
    wchar_t api_key[512] = L"";
    wchar_t key_text[32] = L"R";
    wchar_t mods_text[32] = L"3";
    wchar_t backend_name[32] = L"groq";
    wchar_t sample_rate_text[32] = L"16000";
    wchar_t voice_threshold_text[32] = L"1400";
    wchar_t silence_timeout_text[32] = L"1500";
    wchar_t min_record_text[32] = L"900";
    wchar_t max_record_text[32] = L"30000";
    wchar_t mic_dev_text[32] = L"4294967295";
    wchar_t replace_rules_text[2048] = L"";
    wchar_t continuous_mode_text[16] = L"0";
    wchar_t auto_stop_text[16] = L"1";
    wchar_t translate_text[16] = L"0";
    wchar_t local_model_text[16] = L"0";
    UINT mods = 0;

    if (!app) {
        return;
    }

    app->backend = ASR_BACKEND_GROQ;
    app->continuous_mode = FALSE;
    app->auto_stop_enabled = TRUE;
    app->stop_after_current = FALSE;
    app->translate_enabled = FALSE;
    app->sherpa_exe[0] = L'\0';
    app->sherpa_args[0] = L'\0';
    app->replace_rules[0] = L'\0';
    app->gemini_key[0] = L'\0';
    app->project_id[0] = L'\0';
    app->gemini_model[0] = L'\0';
    app->gladia_key[0] = L'\0';
    app->target_lang[0] = L'\0';
    set_default_recorder_config(app);

    GetPrivateProfileStringW(L"settings", L"api_key", L"", api_key, _countof(api_key), app->config_path);
    GetPrivateProfileStringW(L"settings", L"hotkey_key", L"R", key_text, _countof(key_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"hotkey_mods", L"3", mods_text, _countof(mods_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"backend", L"groq", backend_name, _countof(backend_name), app->config_path);
    GetPrivateProfileStringW(L"settings", L"sherpa_exe", L"", app->sherpa_exe, _countof(app->sherpa_exe), app->config_path);
    GetPrivateProfileStringW(L"settings", L"sherpa_args", L"", app->sherpa_args, _countof(app->sherpa_args), app->config_path);
    GetPrivateProfileStringW(L"settings", L"replace_rules", L"", replace_rules_text, _countof(replace_rules_text), app->config_path);

    GetPrivateProfileStringW(L"settings", L"gemini_key", L"", app->gemini_key, _countof(app->gemini_key), app->config_path);
    GetPrivateProfileStringW(L"settings", L"project_id", L"", app->project_id, _countof(app->project_id), app->config_path);
    GetPrivateProfileStringW(L"settings", L"gemini_model", L"gemini-3.1-flash-lite-preview", app->gemini_model, _countof(app->gemini_model), app->config_path);
    GetPrivateProfileStringW(L"settings", L"gemini_prompt", L"You are an AI text processing assistant.", app->gemini_prompt, _countof(app->gemini_prompt), app->config_path);
    GetPrivateProfileStringW(L"settings", L"thinking_level", L"LOW", app->thinking_level, _countof(app->thinking_level), app->config_path);
    GetPrivateProfileStringW(L"settings", L"gladia_key", L"", app->gladia_key, _countof(app->gladia_key), app->config_path);
    GetPrivateProfileStringW(L"settings", L"target_lang", L"不翻译", app->target_lang, _countof(app->target_lang), app->config_path);
    GetPrivateProfileStringW(L"settings", L"translate_enabled", L"0", translate_text, _countof(translate_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"local_model_index", L"0", local_model_text, _countof(local_model_text), app->config_path);

    GetPrivateProfileStringW(L"settings", L"continuous_mode", L"0", continuous_mode_text, _countof(continuous_mode_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"auto_stop", L"1", auto_stop_text, _countof(auto_stop_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"sample_rate", L"16000", sample_rate_text, _countof(sample_rate_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"voice_threshold", L"1400", voice_threshold_text, _countof(voice_threshold_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"silence_timeout_ms", L"1500", silence_timeout_text, _countof(silence_timeout_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"min_record_ms", L"900", min_record_text, _countof(min_record_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"max_record_ms", L"30000", max_record_text, _countof(max_record_text), app->config_path);
    GetPrivateProfileStringW(L"settings", L"mic_device_id", L"4294967295", mic_dev_text, _countof(mic_dev_text), app->config_path);

    mods = (UINT)wcstoul(mods_text, NULL, 10);
    if (mods == 0) {
        mods = MOD_CONTROL | MOD_ALT;
    }

    app->backend = asr_parse_backend_name(backend_name);
    app->continuous_mode = wcstoul(continuous_mode_text, NULL, 10) ? TRUE : FALSE;
    app->auto_stop_enabled = wcstoul(auto_stop_text, NULL, 10) ? TRUE : FALSE;
    app->translate_enabled = wcstoul(translate_text, NULL, 10) ? TRUE : FALSE;
    app->local_model_index = _wtoi(local_model_text);
    app->mic_device_id = (UINT)wcstoul(mic_dev_text, NULL, 10);
    app->recorder_config.device_id = app->mic_device_id;
    wcsncpy_s(app->replace_rules, _countof(app->replace_rules), replace_rules_text, _TRUNCATE);
    trim_wide_whitespace(app->replace_rules);

    app->recorder_config.sample_rate = clamp_setting((DWORD)wcstoul(sample_rate_text, NULL, 10), 8000, 48000);
    app->recorder_config.channels = 1;
    app->recorder_config.bits_per_sample = 16;
    app->recorder_config.voice_threshold = (SHORT)clamp_setting((DWORD)wcstoul(voice_threshold_text, NULL, 10), 120, 6000);
    app->recorder_config.silence_timeout_ms = clamp_setting((DWORD)wcstoul(silence_timeout_text, NULL, 10), 400, 6000);
    app->recorder_config.min_record_ms = clamp_setting((DWORD)wcstoul(min_record_text, NULL, 10), 300, 5000);
    app->recorder_config.max_record_ms = clamp_setting((DWORD)wcstoul(max_record_text, NULL, 10), 3000, 120000);

    try_auto_fill_sherpa_defaults(app);

    SetWindowTextW(app->api_edit, api_key);
    SetWindowTextW(app->gemini_key_edit, app->gemini_key);
    SetWindowTextW(app->project_id_edit, app->project_id);
    SetWindowTextW(app->gladia_key_edit, app->gladia_key);
    SetWindowTextW(app->hotkey_edit, key_text);
    set_checked(app->check_ctrl, (mods & MOD_CONTROL) != 0);
    set_checked(app->check_alt, (mods & MOD_ALT) != 0);
    set_checked(app->check_shift, (mods & MOD_SHIFT) != 0);
    set_checked(app->check_win, (mods & MOD_WIN) != 0);
    set_checked(app->translate_check, app->translate_enabled);
    sync_runtime_settings_to_ui(app);
    save_settings(app);

    app_log_line(app,
                 "settings loaded backend=%s continuous=%u auto_stop=%u threshold=%d silence=%lu min=%lu max=%lu",
                 app->backend == ASR_BACKEND_SHERPA ? "sherpa" : "groq",
                 app->continuous_mode ? 1u : 0u,
                 app->auto_stop_enabled ? 1u : 0u,
                 (int)app->recorder_config.voice_threshold,
                 (unsigned long)app->recorder_config.silence_timeout_ms,
                 (unsigned long)app->recorder_config.min_record_ms,
                 (unsigned long)app->recorder_config.max_record_ms);
}

static BOOL register_record_hotkey(AppState *app, UINT mods, UINT vk) {
    if (!app) {
        return FALSE;
    }

    if (app->hotkey_registered) {
        UnregisterHotKey(app->main_hwnd, HOTKEY_RECORD_ID);
        app->hotkey_registered = FALSE;
    }

    if (!RegisterHotKey(app->main_hwnd, HOTKEY_RECORD_ID, mods, vk)) {
        return FALSE;
    }

    app->hotkey_mods = mods;
    app->hotkey_vk = vk;
    app->hotkey_registered = TRUE;
    update_hotkey_preview(app);
    return TRUE;
}

static BOOL apply_hotkey_from_ui(AppState *app, BOOL persist) {
    wchar_t key_text[32];
    UINT mods = 0;
    UINT vk = 0;

    if (!app) {
        return FALSE;
    }

    mods = read_modifiers(app);
    if (mods == 0) {
        set_status(app, L"请至少选择一个修饰键（Ctrl/Alt/Shift/Win）。");
        return FALSE;
    }

    GetWindowTextW(app->hotkey_edit, key_text, _countof(key_text));
    vk = parse_hotkey_key(key_text);
    if (vk == 0) {
        set_status(app, L"快捷键无效，请使用 A-Z、0-9、F1-F24、Enter 或 Space。");
        return FALSE;
    }

    if (!register_record_hotkey(app, mods, vk)) {
        set_status(app, L"注册快捷键失败，可能被其它程序占用。");
        return FALSE;
    }

    if (persist) {
        save_settings(app);
    }

    set_status(app, L"快捷键设置成功。");
    return TRUE;
}

static BOOL add_tray_icon(AppState *app) {
    NOTIFYICONDATAW nid;

    if (!app || !app->main_hwnd) {
        return FALSE;
    }

    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = app->main_hwnd;
    nid.uID = TRAY_ICON_ID;
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    nid.uCallbackMessage = WMAPP_TRAYICON;
    nid.hIcon = LoadIconW(NULL, (LPCWSTR)IDI_APPLICATION);
    wcscpy_s(nid.szTip, _countof(nid.szTip), L"语音输入助手");

    if (!Shell_NotifyIconW(NIM_ADD, &nid)) {
        return FALSE;
    }

    nid.uVersion = NOTIFYICON_VERSION_4;
    Shell_NotifyIconW(NIM_SETVERSION, &nid);
    app->tray_added = TRUE;
    return TRUE;
}

static void remove_tray_icon(AppState *app) {
    NOTIFYICONDATAW nid;

    if (!app || !app->tray_added) {
        return;
    }

    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = app->main_hwnd;
    nid.uID = TRAY_ICON_ID;
    Shell_NotifyIconW(NIM_DELETE, &nid);
    app->tray_added = FALSE;
}

static void show_main_window(AppState *app) {
    if (!app || !app->main_hwnd) {
        return;
    }

    ShowWindow(app->main_hwnd, SW_SHOW);
    ShowWindow(app->main_hwnd, SW_RESTORE);
    SetForegroundWindow(app->main_hwnd);
}

static void show_tray_menu(AppState *app) {
    HMENU menu = NULL;
    POINT point;

    if (!app || !app->main_hwnd) {
        return;
    }

    // Win11 终极绝招：强推焦点并处理挂起的消息
    SetForegroundWindow(app->main_hwnd);
    Sleep(10); // 极小延迟等待系统 Shell 响应

    if (!GetCursorPos(&point)) {
        return;
    }

    menu = CreatePopupMenu();
    if (!menu) {
        return;
    }

    AppendMenuW(menu, MF_STRING, ID_TRAY_OPEN, L"显示设置窗口");
    AppendMenuW(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT, L"完全退出程序");
    
    // 使用 TPM_RIGHTBUTTON | TPM_RETURNCMD 确保在 Win11 下最快响应
    UINT selected = TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY, 
                                   point.x, point.y, 0, app->main_hwnd, NULL);
    
    if (selected == ID_TRAY_OPEN) {
        show_main_window(app);
    } else if (selected == ID_TRAY_EXIT) {
        app->exit_requested = TRUE;
        DestroyWindow(app->main_hwnd);
    }
    
    PostMessageW(app->main_hwnd, WM_NULL, 0, 0);
    DestroyMenu(menu);
}

static void update_floating_position(AppState *app) {
    HWND foreground = NULL;
    DWORD thread_id = 0;
    GUITHREADINFO thread_info;
    POINT caret_pos = {0, 0};
    RECT work_area;
    BOOL has_caret = FALSE;
    int x = 0;
    int y = 0;
    const int width = 12;
    const int height = 12;

    if (!app || !app->float_hwnd) {
        return;
    }

    foreground = GetForegroundWindow();
    if (!foreground || is_our_window(app, foreground)) {
        if (IsWindowVisible(app->float_hwnd)) {
            ShowWindow(app->float_hwnd, SW_HIDE);
        }
        return;
    }

    ZeroMemory(&thread_info, sizeof(thread_info));
    thread_info.cbSize = sizeof(thread_info);
    thread_id = GetWindowThreadProcessId(foreground, NULL);

    if (thread_id != 0 && GetGUIThreadInfo(thread_id, &thread_info)) {
        if (thread_info.hwndFocus) {
            if (thread_info.rcCaret.left != 0 || thread_info.rcCaret.top != 0) {
                caret_pos.x = thread_info.rcCaret.left;
                caret_pos.y = thread_info.rcCaret.bottom;
                ClientToScreen(thread_info.hwndFocus, &caret_pos);
                x = caret_pos.x + 5;
                y = caret_pos.y + 5;
                has_caret = TRUE;
            }
        }
    }

    if (!has_caret) {
        GetCursorPos(&caret_pos);
        x = caret_pos.x + 15;
        y = caret_pos.y + 15;
    }

    if (!SystemParametersInfoW(SPI_GETWORKAREA, 0, &work_area, 0)) {
        work_area.left = 0;
        work_area.top = 0;
        work_area.right = GetSystemMetrics(SM_CXSCREEN);
        work_area.bottom = GetSystemMetrics(SM_CYSCREEN);
    }

    if (x + width > work_area.right) x = work_area.right - width;
    if (x < work_area.left) x = work_area.left;
    if (y + height > work_area.bottom) y = work_area.bottom - height;
    if (y < work_area.top) y = work_area.top;

    RECT current_rect;
    BOOL pos_changed = TRUE;
    if (GetWindowRect(app->float_hwnd, &current_rect)) {
        if (current_rect.left == x && current_rect.top == y &&
            (current_rect.right - current_rect.left) == width &&
            (current_rect.bottom - current_rect.top) == height) {
            pos_changed = FALSE;
        }
    }

    if (pos_changed || !IsWindowVisible(app->float_hwnd)) {
        SetWindowPos(app->float_hwnd, HWND_TOPMOST, x, y, width, height, SWP_NOACTIVATE | SWP_SHOWWINDOW);
    }
}

static DWORD WINAPI transcribe_thread_proc(LPVOID param) {
    TranscribeTask *task = (TranscribeTask *)param;
    TranscribeResult *result = NULL;
    HWND notify_hwnd = NULL;

    if (!task) {
        return 0;
    }

    notify_hwnd = task->notify_hwnd;
    result = (TranscribeResult *)calloc(1, sizeof(TranscribeResult));
    if (result) {
        result->target_window = task->target_window;
        if (task->backend == ASR_BACKEND_SHERPA) {
            result->success = sherpa_transcribe_wav_cli(task->wav_path,
                                                        task->sherpa_exe,
                                                        task->sherpa_args,
                                                        &result->text,
                                                        &result->error_text);
        } else if (task->backend == ASR_BACKEND_GEMINI) {
            result->success = gemini_transcribe_wav(task->wav_path,
                                                    task->api_key,
                                                    task->project_id,
                                                    task->model_id,
                                                    task->custom_prompt,
                                                    task->target_lang,
                                                    task->thinking_level,
                                                    &result->text,
                                                    &result->error_text);
        } else {
            result->success = groq_transcribe_wav(task->wav_path,
                                                  task->api_key,
                                                  &result->text,
                                                  &result->error_text);
        }
    }

    if (task->api_key) free(task->api_key);
    if (task->project_id) free(task->project_id);
    if (task->model_id) free(task->model_id);
    if (task->custom_prompt) free(task->custom_prompt);
    if (task->target_lang) free(task->target_lang);
    if (task->thinking_level) free(task->thinking_level);
    free(task);

    if (notify_hwnd) {
        PostMessageW(notify_hwnd, WMAPP_TRANSCRIBE_DONE, 0, (LPARAM)result);
    } else if (result) {
        if (result->text) {
            groq_free_text(result->text);
        }
        if (result->error_text) {
            groq_free_text(result->error_text);
        }
        free(result);
    }

    return 0;
}

static BOOL start_transcribing(AppState *app, HWND target_window) {
    wchar_t api_key_wide[512];
    char *api_key_utf8 = NULL;
    char *project_id_utf8 = NULL;
    char *model_id_utf8 = NULL;
    TranscribeTask *task = NULL;
    HANDLE worker = NULL;

    if (!app) {
        return FALSE;
    }

    if (app->backend == ASR_BACKEND_GROQ) {
        GetWindowTextW(app->api_edit, api_key_wide, _countof(api_key_wide));
        trim_wide_whitespace(api_key_wide);
        if (api_key_wide[0] == L'\0') {
            set_status(app, L"请先填写 Groq API Key。");
            return FALSE;
        }

        api_key_utf8 = wide_to_utf8_alloc(api_key_wide);
        if (!api_key_utf8) {
            set_status(app, L"读取 API Key 失败。");
            return FALSE;
        }
    } else if (app->backend == ASR_BACKEND_GEMINI) {
        if (app->gemini_key[0] == L'\0') {
            set_status(app, L"请先填写 Gemini API Key。");
            return FALSE;
        }
        api_key_utf8 = wide_to_utf8_alloc(app->gemini_key);
        if (app->project_id[0] != L'\0') {
            project_id_utf8 = wide_to_utf8_alloc(app->project_id);
        }
        if (app->gemini_model[0] != L'\0') {
            model_id_utf8 = wide_to_utf8_alloc(app->gemini_model);
        } else {
            model_id_utf8 = wide_to_utf8_alloc(L"gemini-2.5-flash");
        }
    }

    task = (TranscribeTask *)calloc(1, sizeof(TranscribeTask));
    if (!task) {
        free(api_key_utf8);
        free(project_id_utf8);
        free(model_id_utf8);
        set_status(app, L"内存不足。");
        return FALSE;
    }

    task->notify_hwnd = app->main_hwnd;
    wcscpy_s(task->wav_path, _countof(task->wav_path), app->wav_path);
    task->api_key = api_key_utf8;
    task->project_id = project_id_utf8;
    task->model_id = model_id_utf8;
    task->backend = app->backend;
    wcscpy_s(task->sherpa_exe, _countof(task->sherpa_exe), app->sherpa_exe);
    wcscpy_s(task->sherpa_args, _countof(task->sherpa_args), app->sherpa_args);
    task->target_window = target_window;

    task->translate_enabled = app->translate_enabled;
    if (app->translate_enabled) {
        if (app->target_lang[0] != L'\0' && wcscmp(app->target_lang, L"不翻译") != 0) {
            task->target_lang = wide_to_utf8_alloc(app->target_lang);
        }
        if (app->gemini_prompt[0] != L'\0') {
            task->custom_prompt = wide_to_utf8_alloc(app->gemini_prompt);
        }
        if (app->thinking_level[0] != L'\0') {
            task->thinking_level = wide_to_utf8_alloc(app->thinking_level);
        }
    }

    worker = CreateThread(NULL, 0, transcribe_thread_proc, task, 0, NULL);
    if (!worker) {
        free(task->api_key);
        free(task->project_id);
        free(task->model_id);
        free(task);
        set_status(app, L"启动识别线程失败。");
        return FALSE;
    }

    if (app->worker_thread) {
        CloseHandle(app->worker_thread);
    }
    app->worker_thread = worker;
    app->state = VOICE_TRANSCRIBING;
    if (app->backend == ASR_BACKEND_SHERPA) {
        set_status(app, L"本地识别中（Sherpa）...");
    } else if (app->backend == ASR_BACKEND_GEMINI) {
        set_status(app, L"云端识别中（Gemini）...");
    } else {
        set_status(app, L"云端识别中（Groq）...");
    }
    app_log_line(app,
                 "transcribe start backend=%s",
                 app->backend == ASR_BACKEND_SHERPA ? "sherpa" : (app->backend == ASR_BACKEND_GEMINI ? "gemini" : "groq"));
    update_float_button(app);
    return TRUE;
}

static BOOL start_recording_session(AppState *app, HWND target_hint) {
    HWND target = target_hint;
    UINT final_device_id = WAVE_MAPPER;
    wchar_t selected_name[256] = L"";

    if (!app) {
        return FALSE;
    }

    if (!target || is_our_window(app, target)) {
        target = app->follow_target_window;
    }

    // 绝不糊弄方案：根据选中的下拉框名称，在系统中实时反向查找真实的物理 ID
    if (app->mic_combo) {
        LRESULT sel = SendMessageW(app->mic_combo, CB_GETCURSEL, 0, 0);
        if (sel != CB_ERR && sel > 0) {
            SendMessageW(app->mic_combo, CB_GETLBTEXT, (WPARAM)sel, (LPARAM)selected_name);
            
            // 遍历所有物理设备，寻找名称匹配的那一个
            UINT num_devs = waveInGetNumDevs();
            WAVEINCAPSW caps;
            BOOL found = FALSE;
            for (UINT i = 0; i < num_devs; ++i) {
                if (waveInGetDevCapsW(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR) {
                    if (wcscmp(caps.szPname, selected_name) == 0) {
                        final_device_id = i;
                        found = TRUE;
                        break;
                    }
                }
            }
            if (!found) {
                // 如果找不到，极有可能是索引发生了变化，回退到 WAVE_MAPPER 保护程序不崩溃
                final_device_id = WAVE_MAPPER;
            }
        } else {
            final_device_id = WAVE_MAPPER;
        }
    }
    
    app->recorder_config.device_id = final_device_id;

    if (audio_start_recording(&app->recorder_config)) {
        app->target_window = target;
        app->state = VOICE_RECORDING;
        set_status(app, L"录音中...");
        update_float_button(app);
        
        char *mic_name_utf8 = wide_to_utf8_alloc(selected_name[0] ? selected_name : L"WAVE_MAPPER");
        app_log_line(app, "PHYSICAL LOCK recording start target=0x%p device=[%s] phys_id=%u", 
                     target, mic_name_utf8 ? mic_name_utf8 : "unknown", final_device_id);
        free(mic_name_utf8);
        return TRUE;
    }

    set_status(app, L"启动录音失败。");
    app_log_line(app, "recording start failed id=%u", final_device_id);
    return FALSE;
}

static BOOL stop_recording_and_transcribe(AppState *app, BOOL auto_stop) {
    BOOL save_ok = FALSE;
    BOOL should_continue_recording = FALSE;

    if (!app) {
        return FALSE;
    }

    // 只有在自动断句且开启了连续模式时，才保持录音设备开启
    if (auto_stop && app->continuous_mode && !app->stop_after_current && !app->exit_requested && !app->pause_requested) {
        should_continue_recording = TRUE;
    }

    if (should_continue_recording) {
        save_ok = audio_save_chunk_and_continue(app->wav_path);
    } else {
        save_ok = audio_stop_and_save(app->wav_path);
        if (!auto_stop) {
            app->stop_after_current = TRUE;
        }
    }

    if (!save_ok) {
        if (!should_continue_recording) {
            app->state = VOICE_IDLE;
        }
        set_status(app, L"录音保存失败。");
        update_float_button(app);
        app_log_line(app, "recording save failed (continue=%d)", should_continue_recording);
        return FALSE;
    }

    app_log_line(app,
                 "recording segment saved reason=%s continue=%d",
                 auto_stop ? "auto-silence" : "manual",
                 should_continue_recording);

    if (!start_transcribing(app, app->target_window)) {
        if (!should_continue_recording) {
            app->state = VOICE_IDLE;
        }
        update_float_button(app);
        app_log_line(app, "transcribe start failed");
        return FALSE;
    }

    return TRUE;
}

static void poll_recording_runtime(AppState *app) {
    AudioRuntimeStatus status;

    if (!app || app->state != VOICE_RECORDING) {
        return;
    }

    if (!audio_get_runtime_status(&status)) {
        return;
    }

    if (app->float_status) {
        if (status.had_voice) {
            SetWindowTextW(app->float_status, L"录音中(有声音)");
        } else {
            SetWindowTextW(app->float_status, L"录音中(静音)");
        }
    }

    if (!app->auto_stop_enabled) {
        return;
    }

    if (status.should_auto_stop) {
        if (!status.had_voice) {
            audio_abort();
            app->state = VOICE_IDLE;
            update_float_button(app);
            set_status(app, L"未检测到清晰语音，本次未发送。");
            app_log_line(app,
                         "auto-stop dropped(no-voice) duration_ms=%lu silence_ms=%lu peak=%lu bytes=%lu",
                         (unsigned long)status.record_duration_ms,
                         (unsigned long)status.ms_since_voice,
                         (unsigned long)status.peak_level,
                         (unsigned long)status.recorded_bytes);

            if (app->continuous_mode) {
                start_recording_session(app, GetForegroundWindow());
            }
            return;
        }

        set_status(app, L"检测到静音，开始识别...");
        app_log_line(app,
                     "auto-stop duration_ms=%lu silence_ms=%lu peak=%lu bytes=%lu",
                     (unsigned long)status.record_duration_ms,
                     (unsigned long)status.ms_since_voice,
                     (unsigned long)status.peak_level,
                     (unsigned long)status.recorded_bytes);
        stop_recording_and_transcribe(app, TRUE);
    }
}

static void toggle_recording(AppState *app, HWND target_hint) {
    if (!app) {
        return;
    }

    if (app->state == VOICE_TRANSCRIBING) {
        app->stop_after_current = TRUE;
        app->pause_requested = TRUE;
        audio_abort();
        set_status(app, L"已丢弃当前录音：正在等待处理结束并暂停...");
        app_log_line(app, "pause requested while transcribing, recording aborted");
        return;
    }

    if (app->state == VOICE_IDLE || app->state == VOICE_PAUSED) {
        app->stop_after_current = FALSE;
        app->pause_requested = FALSE;
        start_recording_session(app, target_hint);
        return;
    }

    // VOICE_RECORDING 状态下按快捷键，请求暂停
    app->pause_requested = FALSE;
    app->state = VOICE_PAUSED;
    audio_abort();
    if (app->float_status) SetWindowTextW(app->float_status, L"已暂停");
    update_float_button(app);
    set_status(app, L"已暂停录音并丢弃当前语音。");
    app_log_line(app, "recording aborted and dropped due to pause request");
}

static void on_transcribe_done(AppState *app, TranscribeResult *result) {
    BOOL should_restart = FALSE;
    BOOL was_continuing = (app->state == VOICE_TRANSCRIBING && audio_is_recording());

    if (!app) {
        return;
    }

    // 如果之前已经通过 continue 模式保持了录音，则不需要 restart
    if (app->pause_requested) {
        should_restart = FALSE;
    } else {
        should_restart = app->continuous_mode && !app->exit_requested && !app->stop_after_current && !was_continuing;
    }

    if (app->pause_requested) {
        app->state = VOICE_PAUSED;
        app->pause_requested = FALSE;
        if (was_continuing) {
            audio_abort();
        }
    } else {
        app->state = was_continuing ? VOICE_RECORDING : VOICE_IDLE;
    }
    
    if (app->state == VOICE_PAUSED) {
        if (app->float_status) SetWindowTextW(app->float_status, L"已暂停");
    } else if (app->state == VOICE_IDLE) {
        if (app->float_status) SetWindowTextW(app->float_status, L"就绪");
    }

    update_float_button(app);

    if (!result) {
        set_status(app, L"识别失败。");
        return;
    }

    if (!result->success || !result->text) {
        if (result->error_text && result->error_text[0] != '\0') {
            wchar_t *error_wide = utf8_to_wide_alloc(result->error_text);
            if (error_wide) {
                wchar_t short_error[160];
                wcsncpy_s(short_error, _countof(short_error), error_wide, _TRUNCATE);
                set_status(app, short_error);
                free(error_wide);
            } else {
                set_status(app, L"识别失败。");
            }
            app_log_line(app, "transcribe failed: %s", result->error_text);
        } else {
            set_status(app, L"识别失败。");
            app_log_line(app, "transcribe failed with no error text");
        }
    } else {
        trim_ascii_whitespace(result->text);
        if (result->text[0] == '\0') {
            set_status(app, L"未识别到有效语音。");
            app_log_line(app, "transcribe success but empty text");
        } else if (app->state == VOICE_PAUSED) {
            set_status(app, L"已丢弃暂停期间的识别结果。");
            app_log_line(app, "transcribe completely dropped due to pause state");
        } else {
            apply_user_replace_rules(app, &result->text);
            trim_ascii_whitespace(result->text);
            if (result->text[0] == '\0') {
                set_status(app, L"术语替换后文本为空，本次未发送。");
                app_log_line(app, "transcribe dropped after replace rules");
            } else {
                apply_simple_sherpa_punctuation(app, &result->text);
                
                if (result->text && result->text[0] != '\0') {
                    wchar_t *wide_res = utf8_to_wide_alloc(result->text);
                    if (wide_res) {
                        wchar_t ui_msg[1024];
                        swprintf(ui_msg, _countof(ui_msg), L"原文: %ls", wide_res);
                        add_ui_log(app, ui_msg);
                        free(wide_res);
                    }
                }

                if (app->translate_enabled) {
                    char *target_lang_utf8 = NULL;
                    if (app->target_lang[0] != L'\0' && wcscmp(app->target_lang, L"不翻译") != 0) {
                        target_lang_utf8 = wide_to_utf8_alloc(app->target_lang);
                    }
                    
                    char *api_key_utf8 = wide_to_utf8_alloc(app->gemini_key);
                    char *project_id_utf8 = app->project_id[0] != L'\0' ? wide_to_utf8_alloc(app->project_id) : NULL;
                    char *model_id_utf8 = app->gemini_model[0] != L'\0' ? wide_to_utf8_alloc(app->gemini_model) : wide_to_utf8_alloc(L"gemini-2.5-flash");
                    char *custom_prompt_utf8 = app->gemini_prompt[0] != L'\0' ? wide_to_utf8_alloc(app->gemini_prompt) : NULL;
                    char *thinking_level_utf8 = app->thinking_level[0] != L'\0' ? wide_to_utf8_alloc(app->thinking_level) : wide_to_utf8_alloc(L"LOW");
                    char *processed_text = NULL;
                    char *process_error = NULL;

                    if (api_key_utf8 && model_id_utf8) {
                        set_status(app, target_lang_utf8 ? L"Gemini 翻译中..." : L"Gemini 文本处理中...");
                        app_log_line(app, "start processing via gemini (target_lang=%s, thinking=%s)", target_lang_utf8 ? target_lang_utf8 : "none", thinking_level_utf8);
                        if (gemini_process_text(api_key_utf8, project_id_utf8, model_id_utf8, custom_prompt_utf8, target_lang_utf8, thinking_level_utf8, result->text, &processed_text, &process_error)) {
                            app_log_line(app, "gemini process success");
                            free(result->text);
                            result->text = processed_text;
                            
                            if (result->text && result->text[0] != '\0') {
                                wchar_t *wide_proc = utf8_to_wide_alloc(result->text);
                                if (wide_proc) {
                                    wchar_t ui_msg[1024];
                                    swprintf(ui_msg, _countof(ui_msg), L"AI处理: %ls", wide_proc);
                                    add_ui_log(app, ui_msg);
                                    free(wide_proc);
                                }
                            }
                        } else {
                            app_log_line(app, "gemini process failed: %s", process_error ? process_error : "unknown error");
                            // Fallback to original text if processing fails
                        }
                    }

                    if (api_key_utf8) free(api_key_utf8);
                    if (project_id_utf8) free(project_id_utf8);
                    if (model_id_utf8) free(model_id_utf8);
                    if (target_lang_utf8) free(target_lang_utf8);
                    if (custom_prompt_utf8) free(custom_prompt_utf8);
                    if (thinking_level_utf8) free(thinking_level_utf8);
                    if (process_error) free(process_error);
                }

                // 优化：实时获取当前最前台窗口，不再强制跳回录音前的旧窗口
                HWND current_foreground = GetForegroundWindow();
                HWND target = current_foreground;

                // 如果当前前台是我们自己的窗口，则尝试使用之前记录的目标窗口
                if (!target || is_our_window(app, target)) {
                    target = result->target_window ? result->target_window : app->follow_target_window;
                }

                if (target && IsWindow(target) && !is_our_window(app, target)) {
                    SetForegroundWindow(target);
                    Sleep(20); // 稍微等待焦点稳定
                }

                if (!app->pause_requested && app->state != VOICE_PAUSED) {
                    if (injector_paste_utf8(result->text)) {
                        set_status(app, L"已粘贴。");
                        app_log_line(app, "paste success text_len=%u to HWND=0x%p", (unsigned)strlen(result->text), target);
                    } else {
                        set_status(app, L"粘贴失败，请保持光标在目标输入框。");
                        app_log_line(app, "paste failed text_len=%u", (unsigned)strlen(result->text));
                    }
                } else {
                    set_status(app, L"已丢弃暂停期间的识别结果。");
                    app_log_line(app, "paste dropped due to pause text_len=%u", (unsigned)strlen(result->text));
                }
            }
        }
    }

    if (result->text) {
        groq_free_text(result->text);
    }
    if (result->error_text) {
        groq_free_text(result->error_text);
    }
    free(result);

    if (app->worker_thread) {
        CloseHandle(app->worker_thread);
        app->worker_thread = NULL;
    }

    if (app->stop_after_current) {
        app_log_line(app, "stop-after-current consumed");
    }
    app->stop_after_current = FALSE;

    if (should_restart && app->state == VOICE_IDLE) {
        start_recording_session(app, GetForegroundWindow());
    }
}

static void apply_font(HWND control, HFONT font) {
    if (control && font) {
        SendMessageW(control, WM_SETFONT, (WPARAM)font, TRUE);
    }
}

static void refresh_mic_list(AppState *app) {
    UINT num_devs = waveInGetNumDevs();
    WAVEINCAPSW caps;
    UINT i;

    if (!app || !app->mic_combo) return;

    SendMessageW(app->mic_combo, CB_RESETCONTENT, 0, 0);
    SendMessageW(app->mic_combo, CB_ADDSTRING, 0, (LPARAM)L"系统默认录音设备");

    for (i = 0; i < num_devs; ++i) {
        if (waveInGetDevCapsW(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR) {
            SendMessageW(app->mic_combo, CB_ADDSTRING, 0, (LPARAM)caps.szPname);
        }
    }

    if (app->mic_device_id == WAVE_MAPPER) {
        SendMessageW(app->mic_combo, CB_SETCURSEL, 0, 0);
    } else {
        if (app->mic_device_id < num_devs) {
            SendMessageW(app->mic_combo, CB_SETCURSEL, (WPARAM)app->mic_device_id + 1, 0);
        } else {
            SendMessageW(app->mic_combo, CB_SETCURSEL, 0, 0);
            app->mic_device_id = WAVE_MAPPER;
        }
    }
}

static void load_gemini_models_to_combo(AppState *app) {
    wchar_t models_path[MAX_PATH];
    wchar_t app_dir[MAX_PATH];
    FILE *f = NULL;
    char line[256];
    int loaded = 0;

    if (!app || !app->gemini_model_combo) return;

    extract_parent_dir(app->config_path, app_dir, _countof(app_dir));
    swprintf(models_path, _countof(models_path), L"%ls\\gemini_models.txt", app_dir);

    _wfopen_s(&f, models_path, L"rt");
    if (f) {
        while (fgets(line, sizeof(line), f)) {
            wchar_t *wide_line = utf8_to_wide_alloc(line);
            if (wide_line) {
                trim_wide_whitespace(wide_line);
                if (wide_line[0] != L'\0' && wide_line[0] != L'#') {
                    SendMessageW(app->gemini_model_combo, CB_ADDSTRING, 0, (LPARAM)wide_line);
                    loaded++;
                }
                free(wide_line);
            }
        }
        fclose(f);
    } else {
        // If file doesn't exist, write defaults
        _wfopen_s(&f, models_path, L"wt");
        if (f) {
            fputs("# 请在此文件每一行填写一个 Gemini 模型名称\n", f);
            fputs("gemini-2.5-flash\n", f);
            fputs("gemini-3.1-flash-lite-preview\n", f);
            fputs("gemini-2.5-pro\n", f);
            fputs("gemini-2.5-flash-native-audio-preview-12-2025\n", f);
            fclose(f);
        }
        SendMessageW(app->gemini_model_combo, CB_ADDSTRING, 0, (LPARAM)L"gemini-2.5-flash");
        SendMessageW(app->gemini_model_combo, CB_ADDSTRING, 0, (LPARAM)L"gemini-3.1-flash-lite-preview");
        SendMessageW(app->gemini_model_combo, CB_ADDSTRING, 0, (LPARAM)L"gemini-2.5-pro");
        SendMessageW(app->gemini_model_combo, CB_ADDSTRING, 0, (LPARAM)L"gemini-2.5-flash-native-audio-preview-12-2025");
    }
}

static void create_main_controls(AppState *app) {
    HFONT font = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    HWND label = NULL;
    HWND button = NULL;

    label = CreateWindowW(L"STATIC", L"云端 API Key：", WS_CHILD | WS_VISIBLE,
                          20, 16, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->api_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_PASSWORD,
                                    120, 12, 300, 24, app->main_hwnd,
                                    (HMENU)(INT_PTR)IDC_EDIT_API, app->instance, NULL);
    apply_font(app->api_edit, font);

    label = CreateWindowW(L"STATIC", L"录音设备：", WS_CHILD | WS_VISIBLE,
                          440, 16, 80, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->mic_combo = CreateWindowW(L"COMBOBOX", L"",
                                   WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                                   520, 12, 320, 200, app->main_hwnd,
                                   (HMENU)(INT_PTR)IDC_COMBO_MIC, app->instance, NULL);
    apply_font(app->mic_combo, font);

    label = CreateWindowW(L"STATIC", L"识别后端：", WS_CHILD | WS_VISIBLE,
                          20, 50, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->backend_combo = CreateWindowW(L"COMBOBOX", L"",
                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                                       120, 46, 170, 200, app->main_hwnd,
                                       (HMENU)(INT_PTR)IDC_COMBO_BACKEND, app->instance, NULL);
    apply_font(app->backend_combo, font);
    SendMessageW(app->backend_combo, CB_ADDSTRING, 0, (LPARAM)L"Groq (需 Key)");
    SendMessageW(app->backend_combo, CB_ADDSTRING, 0, (LPARAM)L"Sherpa (本地)");
    SendMessageW(app->backend_combo, CB_ADDSTRING, 0, (LPARAM)L"Gladia (需 Key)");
    SendMessageW(app->backend_combo, CB_ADDSTRING, 0, (LPARAM)L"Gemini Native Audio");

    refresh_mic_list(app);

    app->continuous_check = CreateWindowW(L"BUTTON",
                                          L"自动监听（识别完自动继续）",
                                          WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                          330, 48, 260, 20, app->main_hwnd,
                                          (HMENU)(INT_PTR)IDC_CHECK_CONTINUOUS,
                                          app->instance,
                                          NULL);
    apply_font(app->continuous_check, font);

    app->auto_stop_check = CreateWindowW(L"BUTTON",
                                         L"自动静音停录（关闭后需按快捷键结束）",
                                         WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                         590, 48, 250, 20, app->main_hwnd,
                                         (HMENU)(INT_PTR)IDC_CHECK_AUTO_STOP,
                                         app->instance,
                                         NULL);
    apply_font(app->auto_stop_check, font);

    // -- 新增的 Gemini/Gladia 配置区域 (Y = 82 ~ 178) --
    label = CreateWindowW(L"STATIC", L"Gemini Key：", WS_CHILD | WS_VISIBLE,
                          20, 82, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->gemini_key_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                           WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_PASSWORD,
                                           120, 78, 300, 24, app->main_hwnd,
                                           (HMENU)(INT_PTR)IDC_EDIT_GEMINI_KEY, app->instance, NULL);
    apply_font(app->gemini_key_edit, font);

    label = CreateWindowW(L"STATIC", L"Project ID：", WS_CHILD | WS_VISIBLE,
                          440, 82, 80, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->project_id_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                           WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                           520, 78, 200, 24, app->main_hwnd,
                                           (HMENU)(INT_PTR)IDC_EDIT_PROJECT_ID, app->instance, NULL);
    apply_font(app->project_id_edit, font);

    label = CreateWindowW(L"STATIC", L"翻译目标语言：", WS_CHILD | WS_VISIBLE,
                          20, 114, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->lang_combo = CreateWindowW(L"COMBOBOX", L"",
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                                    120, 110, 150, 200, app->main_hwnd,
                                    (HMENU)(INT_PTR)IDC_COMBO_LANG, app->instance, NULL);
    apply_font(app->lang_combo, font);
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"不翻译");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"英语 (English)");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"日语 (Japanese)");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"韩语 (Korean)");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"法语 (French)");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"俄语 (Russian)");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"德语 (German)");
    SendMessageW(app->lang_combo, CB_ADDSTRING, 0, (LPARAM)L"西班牙语 (Spanish)");

    app->translate_check = CreateWindowW(L"BUTTON", L"启用 Gemini 润色/处理/翻译",
                                         WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                         280, 112, 160, 20, app->main_hwnd,
                                         (HMENU)(INT_PTR)IDC_CHECK_TRANSLATE, app->instance, NULL);
    apply_font(app->translate_check, font);

    label = CreateWindowW(L"STATIC", L"Gemini 模型：", WS_CHILD | WS_VISIBLE,
                          440, 114, 80, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->gemini_model_combo = CreateWindowW(L"COMBOBOX", L"",
                                            WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWN | CBS_AUTOHSCROLL,
                                            520, 110, 320, 200, app->main_hwnd,
                                            (HMENU)(INT_PTR)IDC_COMBO_GEMINI_MODEL, app->instance, NULL);
    apply_font(app->gemini_model_combo, font);
    load_gemini_models_to_combo(app);

    label = CreateWindowW(L"STATIC", L"Gladia Key：", WS_CHILD | WS_VISIBLE,
                          20, 146, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->gladia_key_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                           WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_PASSWORD,
                                           120, 142, 300, 24, app->main_hwnd,
                                           (HMENU)(INT_PTR)IDC_EDIT_GLADIA_KEY, app->instance, NULL);
    apply_font(app->gladia_key_edit, font);

    label = CreateWindowW(L"STATIC", L"思考模式(3.1+)：", WS_CHILD | WS_VISIBLE,
                          440, 146, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->thinking_combo = CreateWindowW(L"COMBOBOX", L"",
                                        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                                        540, 142, 100, 200, app->main_hwnd,
                                        (HMENU)(INT_PTR)IDC_COMBO_THINKING, app->instance, NULL);
    apply_font(app->thinking_combo, font);
    SendMessageW(app->thinking_combo, CB_ADDSTRING, 0, (LPARAM)L"NONE");
    SendMessageW(app->thinking_combo, CB_ADDSTRING, 0, (LPARAM)L"LOW");
    SendMessageW(app->thinking_combo, CB_ADDSTRING, 0, (LPARAM)L"MEDIUM");
    SendMessageW(app->thinking_combo, CB_ADDSTRING, 0, (LPARAM)L"HIGH");

    label = CreateWindowW(L"STATIC", L"Gemini 指令：", WS_CHILD | WS_VISIBLE,
                          20, 178, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->gemini_prompt_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                           WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN,
                                           120, 174, 720, 60, app->main_hwnd,
                                           (HMENU)(INT_PTR)IDC_EDIT_GEMINI_PROMPT, app->instance, NULL);
    apply_font(app->gemini_prompt_edit, font);
    // -- 结束 --

    label = CreateWindowW(L"STATIC", L"本地模型：", WS_CHILD | WS_VISIBLE,
                          20, 244, 100, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->model_combo = CreateWindowW(L"COMBOBOX", L"",
                                        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                                        120, 240, 240, 200, app->main_hwnd,
                                        (HMENU)(INT_PTR)IDC_COMBO_MODEL, app->instance, NULL);
    apply_font(app->model_combo, font);
    SendMessageW(app->model_combo, CB_ADDSTRING, 0, (LPARAM)L"默认模型 (Paraformer)");
    SendMessageW(app->model_combo, CB_ADDSTRING, 0, (LPARAM)L"Zipformer-zh-xlarge (2025)");
    SendMessageW(app->model_combo, CB_ADDSTRING, 0, (LPARAM)L"FunASR-nano (2025)");
    SendMessageW(app->model_combo, CB_SETCURSEL, 0, 0);

    HWND btn_apply_model = CreateWindowW(L"BUTTON", L"下载并配置", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                          370, 240, 150, 24, app->main_hwnd, (HMENU)(INT_PTR)IDC_BTN_APPLY_MODEL, app->instance, NULL);
    apply_font(btn_apply_model, font);

    HWND btn_open_config = CreateWindowW(L"BUTTON", L"打开配置目录", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                          530, 240, 150, 24, app->main_hwnd, (HMENU)(INT_PTR)IDC_BTN_OPEN_CONFIG, app->instance, NULL);
    apply_font(btn_open_config, font);

    label = CreateWindowW(L"STATIC", L"Sherpa 程序：", WS_CHILD | WS_VISIBLE,
                          20, 274, 110, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->sherpa_exe_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                           L"EDIT",
                                           L"",
                                           WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                           140,
                                           270,
                                           640,
                                           24,
                                           app->main_hwnd,
                                           (HMENU)(INT_PTR)IDC_EDIT_SHERPA_EXE,
                                           app->instance,
                                           NULL);
    SendMessageW(app->sherpa_exe_edit, EM_SETLIMITTEXT, 2048, 0);
    apply_font(app->sherpa_exe_edit, font);

    label = CreateWindowW(L"STATIC", L"Sherpa 参数：", WS_CHILD | WS_VISIBLE,
                          20, 306, 110, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->sherpa_args_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                            L"EDIT",
                                            L"",
                                            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                            140, 302, 640, 24,
                                            app->main_hwnd,
                                            (HMENU)(INT_PTR)IDC_EDIT_SHERPA_ARGS,
                                            app->instance,
                                            NULL);
    SendMessageW(app->sherpa_args_edit, EM_SETLIMITTEXT, 4096, 0);
    SendMessageW(app->sherpa_args_edit, EM_SETLIMITTEXT, 4096, 0);
    apply_font(app->sherpa_args_edit, font);

    label = CreateWindowW(L"STATIC", L"静音判停参数：", WS_CHILD | WS_VISIBLE,
                          20, 342, 120, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    label = CreateWindowW(L"STATIC", L"音量阈值", WS_CHILD | WS_VISIBLE,
                          140, 342, 70, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->threshold_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                          L"EDIT",
                                          L"1400",
                                          WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                          250, 338, 80, 24,
                                          app->main_hwnd,
                                          (HMENU)(INT_PTR)IDC_EDIT_THRESHOLD,
                                          app->instance,
                                          NULL);
    apply_font(app->threshold_edit, font);

    label = CreateWindowW(L"STATIC", L"静音时长(ms)", WS_CHILD | WS_VISIBLE,
                          310, 342, 90, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->silence_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                        L"EDIT",
                                        L"1500",
                                        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                        400, 338, 80, 24,
                                        app->main_hwnd,
                                        (HMENU)(INT_PTR)IDC_EDIT_SILENCE,
                                        app->instance,
                                        NULL);
    apply_font(app->silence_edit, font);

    label = CreateWindowW(L"STATIC", L"最短录音(ms)", WS_CHILD | WS_VISIBLE,
                          500, 342, 90, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->minrec_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                       L"EDIT",
                                       L"900",
                                       WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                       590, 338, 80, 24,
                                       app->main_hwnd,
                                       (HMENU)(INT_PTR)IDC_EDIT_MINREC,
                                       app->instance,
                                       NULL);
    apply_font(app->minrec_edit, font);

    label = CreateWindowW(L"STATIC", L"最长录音(ms)", WS_CHILD | WS_VISIBLE,
                          680, 342, 90, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->maxrec_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                       L"EDIT",
                                       L"30000",
                                       WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                       770, 338, 70, 24,
                                       app->main_hwnd,
                                       (HMENU)(INT_PTR)IDC_EDIT_MAXREC,
                                       app->instance,
                                       NULL);
    apply_font(app->maxrec_edit, font);

    label = CreateWindowW(L"STATIC", L"快捷键（A-Z/0-9/F1-F24）：", WS_CHILD | WS_VISIBLE,
                          20, 380, 220, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->hotkey_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"R",
                                       WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                       250, 376, 80, 24, app->main_hwnd,
                                       (HMENU)(INT_PTR)IDC_EDIT_HOTKEY, app->instance, NULL);
    apply_font(app->hotkey_edit, font);

    app->check_ctrl = CreateWindowW(L"BUTTON", L"Ctrl", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                    20, 412, 70, 20, app->main_hwnd,
                                    (HMENU)(INT_PTR)IDC_CHECK_CTRL, app->instance, NULL);
    apply_font(app->check_ctrl, font);

    app->check_alt = CreateWindowW(L"BUTTON", L"Alt", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                   95, 412, 70, 20, app->main_hwnd,
                                   (HMENU)(INT_PTR)IDC_CHECK_ALT, app->instance, NULL);
    apply_font(app->check_alt, font);

    app->check_shift = CreateWindowW(L"BUTTON", L"Shift", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                     170, 412, 70, 20, app->main_hwnd,
                                     (HMENU)(INT_PTR)IDC_CHECK_SHIFT, app->instance, NULL);
    apply_font(app->check_shift, font);

    app->check_win = CreateWindowW(L"BUTTON", L"Win", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                   245, 412, 70, 20, app->main_hwnd,
                                   (HMENU)(INT_PTR)IDC_CHECK_WIN, app->instance, NULL);
    apply_font(app->check_win, font);

    app->current_hotkey_label = CreateWindowW(L"STATIC", L"当前快捷键：未设置",
                                              WS_CHILD | WS_VISIBLE,
                                              20, 440, 260, 20, app->main_hwnd,
                                              (HMENU)(INT_PTR)IDC_LABEL_CURRENT_HOTKEY,
                                              app->instance, NULL);
    apply_font(app->current_hotkey_label, font);

    button = CreateWindowW(L"BUTTON", L"保存设置",
                           WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                           530, 376, 120, 30, app->main_hwnd,
                           (HMENU)(INT_PTR)IDC_BTN_APPLY, app->instance, NULL);
    apply_font(button, font);

    button = CreateWindowW(L"BUTTON", L"配置自检",
                           WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                           660, 376, 120, 30, app->main_hwnd,
                           (HMENU)(INT_PTR)IDC_BTN_SELF_CHECK, app->instance, NULL);
    apply_font(button, font);

    button = CreateWindowW(L"BUTTON", L"安装本地模型（Sherpa）",
                           WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                           530, 412, 250, 30, app->main_hwnd,
                           (HMENU)(INT_PTR)IDC_BTN_INSTALL_SHERPA, app->instance, NULL);
    apply_font(button, font);

    button = CreateWindowW(L"BUTTON", L"退出程序",
                           WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                           530, 448, 250, 24, app->main_hwnd,
                           (HMENU)(INT_PTR)IDC_BTN_EXIT, app->instance, NULL);
    apply_font(button, font);

    label = CreateWindowW(L"STATIC", L"术语纠错（错词=正词；多条用分号）：", WS_CHILD | WS_VISIBLE,
                          20, 466, 260, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->replace_rules_edit = CreateWindowExW(WS_EX_CLIENTEDGE,
                                              L"EDIT",
                                              L"",
                                              WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                              280, 462, 240, 24,
                                              app->main_hwnd,
                                              (HMENU)(INT_PTR)IDC_EDIT_REPLACE_RULES,
                                              app->instance,
                                              NULL);
    apply_font(app->replace_rules_edit, font);

    app->selfcheck_label = CreateWindowW(L"STATIC",
                                         L"尚未执行自检。",
                                         WS_CHILD | WS_VISIBLE | WS_BORDER,
                                         20, 498, 820, 64,
                                         app->main_hwnd,
                                         (HMENU)(INT_PTR)IDC_LABEL_SELFCHECK,
                                         app->instance,
                                         NULL);
    apply_font(app->selfcheck_label, font);

    app->status_label = CreateWindowW(L"STATIC", L"就绪。",
                                      WS_CHILD | WS_VISIBLE,
                                      20, 572, 820, 20, app->main_hwnd,
                                      (HMENU)(INT_PTR)IDC_LABEL_STATUS, app->instance, NULL);
    apply_font(app->status_label, font);

    label = CreateWindowW(L"STATIC",
                          L"关闭窗口只会最小化到托盘；若要完全退出，请点击“退出程序”或托盘菜单“退出程序”。",
                          WS_CHILD | WS_VISIBLE,
                          20, 600, 820, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    label = CreateWindowW(L"STATIC",
                          L"说明：静音时长(ms)越小越容易触发自动停止；最长录音(ms)到达后会强制结束。",
                          WS_CHILD | WS_VISIBLE,
                          20, 622, 820, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    label = CreateWindowW(L"STATIC", L"识别与处理日志 (仅保留最新100条)：", WS_CHILD | WS_VISIBLE,
                          20, 650, 400, 20, app->main_hwnd, NULL, app->instance, NULL);
    apply_font(label, font);

    app->log_list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                                    WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_HASSTRINGS | LBS_NOTIFY,
                                    20, 670, 820, 150, app->main_hwnd,
                                    (HMENU)(INT_PTR)IDC_LIST_LOG, app->instance, NULL);
    apply_font(app->log_list, font);
}
static LRESULT CALLBACK FloatWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    AppState *app = (AppState *)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    switch (msg) {
    case WM_NCCREATE: {
        CREATESTRUCTW *create = (CREATESTRUCTW *)lParam;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)create->lpCreateParams);
        return TRUE;
    }
    case WM_CREATE:
        app = (AppState *)((CREATESTRUCTW *)lParam)->lpCreateParams;
        if (!app) {
            return -1;
        }
        app->float_hwnd = hwnd;
        SetLayeredWindowAttributes(hwnd, RGB(255, 0, 255), 200, LWA_COLORKEY | LWA_ALPHA);
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);
        
        HBRUSH bgBrush = CreateSolidBrush(RGB(255, 0, 255));
        FillRect(hdc, &rc, bgBrush);
        DeleteObject(bgBrush);

        COLORREF color = RGB(0, 120, 255); // Blue
        if (app) {
            if (app->state == VOICE_RECORDING || app->state == VOICE_TRANSCRIBING) {
                color = RGB(0, 220, 0); // Green
            } else if (app->state == VOICE_PAUSED) {
                color = RGB(255, 200, 0); // Yellow
            }
        }

        HBRUSH brush = CreateSolidBrush(color);
        HPEN pen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
        HGDIOBJ oldBrush = SelectObject(hdc, brush);
        HGDIOBJ oldPen = SelectObject(hdc, pen);

        Ellipse(hdc, rc.left, rc.top, rc.right, rc.bottom);

        SelectObject(hdc, oldBrush);
        SelectObject(hdc, oldPen);
        DeleteObject(brush);
        DeleteObject(pen);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_NCHITTEST:
        return HTTRANSPARENT;
    case WM_CLOSE:
        ShowWindow(hwnd, SW_HIDE);
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    AppState *app = (AppState *)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    switch (msg) {
    case WM_NCCREATE: {
        CREATESTRUCTW *create = (CREATESTRUCTW *)lParam;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)create->lpCreateParams);
        return TRUE;
    }
    case WM_CREATE:
        app = (AppState *)((CREATESTRUCTW *)lParam)->lpCreateParams;
        if (!app) {
            return -1;
        }
        app->main_hwnd = hwnd;
        create_main_controls(app);
        return 0;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_CHECK_CONTINUOUS:
            if (HIWORD(wParam) == BN_CLICKED) {
                app->continuous_mode = is_checked(app->continuous_check);
                app_log_line(app, "continuous_mode toggled to %d", app->continuous_mode);
                save_settings(app);
            }
            return 0;
        case IDC_CHECK_AUTO_STOP:
            if (HIWORD(wParam) == BN_CLICKED) {
                app->auto_stop_enabled = is_checked(app->auto_stop_check);
                app_log_line(app, "auto_stop toggled to %d", app->auto_stop_enabled);
                save_settings(app);
            }
            return 0;
        case IDC_COMBO_MIC:
            if (HIWORD(wParam) == CBN_SELCHANGE) {
                LRESULT sel = SendMessageW(app->mic_combo, CB_GETCURSEL, 0, 0);
                if (sel == 0) {
                    app->mic_device_id = WAVE_MAPPER;
                } else {
                    app->mic_device_id = (UINT)(sel - 1);
                }
                app->recorder_config.device_id = app->mic_device_id;
                app_log_line(app, "mic_device_id changed to %u", app->mic_device_id);
                save_settings(app);
            }
            return 0;
        case IDC_BTN_APPLY:
            if (apply_hotkey_from_ui(app, FALSE)) {
                apply_runtime_settings_from_ui(app, TRUE);
                run_self_check(app, FALSE);
                set_status(app, L"设置已保存。");
            }
            return 0;
        case IDC_BTN_SELF_CHECK:
            run_self_check(app, TRUE);
            return 0;
        case IDC_COMBO_MODEL:
            if (HIWORD(wParam) == CBN_SELCHANGE) {
                LRESULT sel = SendMessageW(app->model_combo, CB_GETCURSEL, 0, 0);
                apply_model_selection(app, (int)sel);
            }
            return 0;
        case IDC_BTN_APPLY_MODEL: {
            LRESULT sel = SendMessageW(app->model_combo, CB_GETCURSEL, 0, 0);
            apply_model_selection(app, (int)sel);
            return 0;
        }
        case IDC_BTN_OPEN_CONFIG: {
            wchar_t config_dir[MAX_PATH];
            extract_parent_dir(app->config_path, config_dir, _countof(config_dir));
            ShellExecuteW(NULL, L"open", config_dir, NULL, NULL, SW_SHOWNORMAL);
            return 0;
        }
        case IDC_BTN_INSTALL_SHERPA:
            launch_sherpa_installer(app);
            return 0;
        case IDC_BTN_EXIT:
            app->exit_requested = TRUE;
            DestroyWindow(hwnd);
            return 0;
        case ID_TRAY_OPEN:
            show_main_window(app);
            return 0;
        case ID_TRAY_EXIT:
            app->exit_requested = TRUE;
            DestroyWindow(hwnd);
            return 0;
        default:
            break;
        }
        break;
    case WM_HOTKEY:
        if (wParam == HOTKEY_RECORD_ID) {
            HWND target = GetForegroundWindow();
            if (!target || is_our_window(app, target)) {
                target = app->follow_target_window;
            }
            toggle_recording(app, target);
            return 0;
        }
        break;
    case WM_TIMER:
        if (wParam == TIMER_FOLLOW_INPUT) {
            update_floating_position(app);
            poll_recording_runtime(app);
            return 0;
        }
        break;
    case WMAPP_FLOAT_TOGGLE:
        toggle_recording(app, app->follow_target_window);
        return 0;
    case WMAPP_TRANSCRIBE_DONE:
        on_transcribe_done(app, (TranscribeResult *)lParam);
        return 0;
    case WMAPP_TRAYICON:
        {
            UINT tray_msg = LOWORD(lParam);
            // 增强：捕获所有可能的右键点击消息
            if (tray_msg == WM_CONTEXTMENU || tray_msg == WM_RBUTTONUP || tray_msg == WM_RBUTTONDOWN) {
                show_tray_menu(app);
                return 0;
            }
            if (tray_msg == NIN_SELECT || tray_msg == WM_LBUTTONUP || tray_msg == WM_LBUTTONDBLCLK) {
                show_main_window(app);
                return 0;
            }
        }
        break;
    case WM_CLOSE:
        if (app && !app->exit_requested) {
            ShowWindow(hwnd, SW_HIDE);
            set_status(app, L"窗口已隐藏到托盘，程序仍在运行。");
            return 0;
        }
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        if (app) {
            KillTimer(hwnd, TIMER_FOLLOW_INPUT);
            audio_abort();

            if (app->hotkey_registered) {
                UnregisterHotKey(app->main_hwnd, HOTKEY_RECORD_ID);
                app->hotkey_registered = FALSE;
            }

            if (app->worker_thread) {
                WaitForSingleObject(app->worker_thread, 5000);
                CloseHandle(app->worker_thread);
                app->worker_thread = NULL;
            }

            if (app->float_hwnd) {
                DestroyWindow(app->float_hwnd);
                app->float_hwnd = NULL;
            }

            remove_tray_icon(app);
            DeleteFileW(app->wav_path);
        }

        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    AppState app;
    WNDCLASSEXW main_class;
    WNDCLASSEXW float_class;
    MSG message;

    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    ZeroMemory(&app, sizeof(app));
    app.instance = hInstance;
    app.state = VOICE_IDLE;

    if (!build_temp_wav_path(app.wav_path, _countof(app.wav_path))) {
        MessageBoxW(NULL, L"无法创建临时录音文件路径。", APP_TITLE, MB_ICONERROR);
        return 1;
    }

    if (!build_config_path(app.config_path, _countof(app.config_path))) {
        MessageBoxW(NULL, L"无法创建配置文件路径。", APP_TITLE, MB_ICONERROR);
        return 1;
    }

    if (!build_log_path(app.config_path, app.log_path, _countof(app.log_path))) {
        app.log_path[0] = L'\0';
    } else {
        HANDLE hLog = CreateFileW(app.log_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hLog != INVALID_HANDLE_VALUE) {
            CloseHandle(hLog);
        }
    }

    ZeroMemory(&main_class, sizeof(main_class));
    main_class.cbSize = sizeof(main_class);
    main_class.lpfnWndProc = MainWndProc;
    main_class.hInstance = hInstance;
    main_class.hCursor = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    main_class.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    main_class.lpszClassName = MAIN_CLASS_NAME;

    if (!RegisterClassExW(&main_class)) {
        return 1;
    }

    ZeroMemory(&float_class, sizeof(float_class));
    float_class.cbSize = sizeof(float_class);
    float_class.lpfnWndProc = FloatWndProc;
    float_class.hInstance = hInstance;
    float_class.hCursor = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    float_class.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    float_class.lpszClassName = FLOAT_CLASS_NAME;

    if (!RegisterClassExW(&float_class)) {
        return 1;
    }

    app.main_hwnd = CreateWindowExW(0, MAIN_CLASS_NAME, APP_TITLE,
                                    WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                                    CW_USEDEFAULT, CW_USEDEFAULT, 920, 860,
                                    NULL, NULL, hInstance, &app);
    if (!app.main_hwnd) {
        return 1;
    }

    app.float_hwnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_TRANSPARENT, FLOAT_CLASS_NAME, L"语音输入悬浮窗",
                                     WS_POPUP,
                                     CW_USEDEFAULT, CW_USEDEFAULT, 12, 12,
                                     NULL, NULL, hInstance, &app);
    if (!app.float_hwnd) {
        DestroyWindow(app.main_hwnd);
        return 1;
    }

    load_settings(&app);
    app_log_line(&app, "application startup");
    if (!apply_hotkey_from_ui(&app, FALSE)) {
        SetWindowTextW(app.hotkey_edit, L"R");
        set_checked(app.check_ctrl, TRUE);
        set_checked(app.check_alt, TRUE);
        set_checked(app.check_shift, FALSE);
        set_checked(app.check_win, FALSE);
        apply_hotkey_from_ui(&app, FALSE);
    }

    if (!add_tray_icon(&app)) {
        MessageBoxW(app.main_hwnd, L"托盘图标添加失败。", APP_TITLE, MB_ICONWARNING);
    }

    SetTimer(app.main_hwnd, TIMER_FOLLOW_INPUT, 16, NULL);
    run_self_check(&app, FALSE);
    update_float_button(&app);

    ShowWindow(app.main_hwnd, SW_SHOW);
    UpdateWindow(app.main_hwnd);
    ShowWindow(app.float_hwnd, SW_HIDE);

    while (GetMessageW(&message, NULL, 0, 0) > 0) {
        if (message.message == WM_KEYDOWN && message.wParam == 'A' && (GetKeyState(VK_CONTROL) < 0)) {
            HWND hFocus = GetFocus();
            if (hFocus) {
                wchar_t className[32];
                if (GetClassNameW(hFocus, className, 32)) {
                    if (_wcsicmp(className, L"Edit") == 0) {
                        SendMessageW(hFocus, EM_SETSEL, 0, -1);
                        continue;
                    }
                }
            }
        }
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }

    return (int)message.wParam;
}

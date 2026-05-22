#include "input_injector.h"

#include <stdlib.h>
#include <string.h>

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

static BOOL send_ctrl_v(void) {
    INPUT inputs[4];

    ZeroMemory(inputs, sizeof(inputs));

    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_CONTROL;

    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = 'V';

    inputs[2].type = INPUT_KEYBOARD;
    inputs[2].ki.wVk = 'V';
    inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[3].type = INPUT_KEYBOARD;
    inputs[3].ki.wVk = VK_CONTROL;
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

    return SendInput(4, inputs, sizeof(INPUT)) == 4;
}

BOOL injector_paste_utf8(const char *utf8_text) {
    wchar_t *wide_text = NULL;
    size_t len = 0;
    INPUT *inputs = NULL;
    size_t i = 0;
    UINT sent = 0;

    if (!utf8_text || utf8_text[0] == '\0') {
        return FALSE;
    }

    wide_text = utf8_to_wide(utf8_text);
    if (!wide_text) {
        return FALSE;
    }

    len = wcslen(wide_text);
    if (len == 0) {
        free(wide_text);
        return FALSE;
    }

    inputs = (INPUT *)calloc(len * 2, sizeof(INPUT));
    if (!inputs) {
        free(wide_text);
        return FALSE;
    }

    for (i = 0; i < len; ++i) {
        wchar_t wch = wide_text[i];

        // Key down
        inputs[i * 2].type = INPUT_KEYBOARD;
        inputs[i * 2].ki.wVk = 0;
        inputs[i * 2].ki.wScan = wch;
        inputs[i * 2].ki.dwFlags = KEYEVENTF_UNICODE;

        // Key up
        inputs[i * 2 + 1].type = INPUT_KEYBOARD;
        inputs[i * 2 + 1].ki.wVk = 0;
        inputs[i * 2 + 1].ki.wScan = wch;
        inputs[i * 2 + 1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
    }

    sent = SendInput((UINT)(len * 2), inputs, sizeof(INPUT));

    free(inputs);
    free(wide_text);

    return sent == (UINT)(len * 2);
}

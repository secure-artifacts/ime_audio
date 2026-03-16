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
    size_t bytes = 0;
    HGLOBAL clipboard_mem = NULL;
    wchar_t *mem_ptr = NULL;

    if (!utf8_text || utf8_text[0] == '\0') {
        return FALSE;
    }

    wide_text = utf8_to_wide(utf8_text);
    if (!wide_text) {
        return FALSE;
    }

    bytes = (wcslen(wide_text) + 1) * sizeof(wchar_t);
    clipboard_mem = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!clipboard_mem) {
        free(wide_text);
        return FALSE;
    }

    mem_ptr = (wchar_t *)GlobalLock(clipboard_mem);
    if (!mem_ptr) {
        GlobalFree(clipboard_mem);
        free(wide_text);
        return FALSE;
    }

    memcpy(mem_ptr, wide_text, bytes);
    GlobalUnlock(clipboard_mem);
    free(wide_text);

    if (!OpenClipboard(NULL)) {
        GlobalFree(clipboard_mem);
        return FALSE;
    }

    EmptyClipboard();
    if (!SetClipboardData(CF_UNICODETEXT, clipboard_mem)) {
        CloseClipboard();
        GlobalFree(clipboard_mem);
        return FALSE;
    }

    CloseClipboard();
    Sleep(40);
    return send_ctrl_v();
}

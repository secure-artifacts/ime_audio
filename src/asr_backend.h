#ifndef ASR_BACKEND_H
#define ASR_BACKEND_H

#include <windows.h>

typedef enum AsrBackendKind {
    ASR_BACKEND_GROQ = 0,
    ASR_BACKEND_SHERPA = 1,
    ASR_BACKEND_GLADIA = 2,
    ASR_BACKEND_GEMINI = 3
} AsrBackendKind;

AsrBackendKind asr_parse_backend_name(const wchar_t *name);
const wchar_t *asr_backend_name(AsrBackendKind backend);

BOOL sherpa_transcribe_wav_cli(const wchar_t *wav_path,
                               const wchar_t *sherpa_exe,
                               const wchar_t *sherpa_args,
                               char **out_utf8_text,
                               char **out_error_utf8);

#endif

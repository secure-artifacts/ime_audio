#ifndef GEMINI_CLIENT_H
#define GEMINI_CLIENT_H

#include <windows.h>
#include <stdbool.h>

BOOL gemini_process_text(const char *api_key,
                         const char *project_id,
                         const char *model_id,
                         const char *custom_prompt,
                         const char *target_lang,
                         const char *thinking_level,
                         const char *input_text,
                         char **out_text_utf8,
                         char **out_error_utf8);

BOOL gemini_transcribe_wav(const wchar_t *wav_path,
                           const char *api_key,
                           const char *project_id,
                           const char *model_id,
                           const char *custom_prompt,
                           const char *target_lang,
                           const char *thinking_level,
                           char **out_utf8_text,
                           char **out_error_utf8);

void gemini_free_text(char *text);

#endif

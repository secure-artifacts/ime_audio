#ifndef GROQ_CLIENT_H
#define GROQ_CLIENT_H

#include <windows.h>

BOOL groq_transcribe_wav(const wchar_t *wav_path,
						 const char *api_key,
						 char **out_utf8_text,
						 char **out_error_utf8);

BOOL gladia_transcribe_wav(const wchar_t *wav_path,
                           const char *api_key,
                           char **out_utf8_text,
                           char **out_error_utf8);

void groq_free_text(char *text);

#endif


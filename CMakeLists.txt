cmake_minimum_required(VERSION 3.20)
project(voice_ime_c LANGUAGES C)

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)

add_executable(voice_ime WIN32
    src/main.c
    src/audio_recorder.c
    src/asr_backend.c
    src/groq_client.c
    src/input_injector.c
)

target_compile_definitions(voice_ime PRIVATE _CRT_SECURE_NO_WARNINGS)
if (MSVC)
    target_compile_options(voice_ime PRIVATE /utf-8)
endif()
target_link_libraries(voice_ime PRIVATE winmm winhttp user32 shell32)

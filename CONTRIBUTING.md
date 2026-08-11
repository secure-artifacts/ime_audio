# Contributing to VoiceIME

Thank you for helping improve VoiceIME. Bug reports, focused feature proposals, documentation updates, tests, and code fixes are welcome.

## Before you start

- Search existing issues and pull requests to avoid duplicates.
- Use the provided issue template and include the smallest reproducible example you can share safely.
- For substantial behavior or architecture changes, open a feature request before investing in an implementation.
- Report vulnerabilities through [SECURITY.md](SECURITY.md), not through a public issue.
- Never attach real API keys, `voice_ime.ini`, private recordings, unredacted logs, or sensitive transcriptions.

## Build locally

Requirements:

- Windows 10 or 11
- Visual Studio 2022 with the Desktop development with C++ workload
- CMake 3.20 or newer

Configure and build:

```powershell
cmake -S . -B build -G "Visual Studio 17 2022"
cmake --build build --config Release
```

The executable is written to `build\Release\voice_ime.exe`.

## Pull request checklist

- Keep the change focused and explain its user-visible effect.
- Describe how the change was tested, including the Windows version and selected backend when relevant.
- Preserve compatibility with existing `voice_ime.ini`, `terms.tsv`, and prompt-template formats, or document the migration.
- Check all buffer sizes, string encodings, paths, process arguments, network responses, and cleanup paths in native-code changes.
- Add or update documentation for user-facing settings and workflows.
- Do not modify release versions, create tags, or publish artifacts unless requested by the maintainer.
- Confirm that no generated binaries, models, credentials, private logs, or local configuration files are included.

## Commit and review guidance

Use clear, imperative commit messages such as `fix: reject an invalid custom endpoint`. A pull request may be revised for correctness, security, compatibility, scope, or maintainability before it is accepted.

By contributing, you agree that your contribution is provided under the repository's [MIT License](LICENSE).

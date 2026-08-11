# Security Policy

VoiceIME operates at a sensitive boundary between microphone input, third-party speech services, local model runtimes, and system-wide text injection. Security reports are taken seriously.

## Supported versions

Security fixes are provided for the latest published release. Before reporting a problem, confirm that it is reproducible on the latest version available from [GitHub Releases](https://github.com/secure-artifacts/ime_audio/releases/latest).

## Reporting a vulnerability

Please do **not** disclose a suspected vulnerability in a public issue, discussion, pull request, log, or screenshot.

1. Use [GitHub private vulnerability reporting](https://github.com/secure-artifacts/ime_audio/security/advisories/new) when it is available.
2. If the private reporting form is unavailable, open a public issue that contains only a request to establish a private contact channel. Do not include technical details, proof-of-concept code, secrets, private audio, or user data.
3. Include the affected version, Windows version, impact, reproducible steps, and any suggested mitigation in the private report.

The maintainer will acknowledge a complete report as soon as practical, investigate it, and coordinate remediation and disclosure. Please allow reasonable time for a fix before publishing details.

## Security-sensitive areas

Reports are especially useful for:

- exposure or unsafe storage of API keys and authentication headers;
- unintended capture, retention, logging, or transmission of microphone audio or transcribed text;
- memory-safety errors in native C/Win32 code;
- command, path, clipboard, keystroke, or prompt injection;
- unsafe handling of custom API endpoints or HTTP/TLS responses;
- untrusted model, runtime, or installer downloads;
- GitHub Actions, OIDC, artifact-attestation, installer, or release supply-chain weaknesses;
- privilege-boundary problems involving global hotkeys, child processes, or text injection.

## Protecting sensitive data

The local `voice_ime.ini` file may contain service credentials. Diagnostic logs can reveal system details or transcribed content. Before sharing any diagnostic material, remove API keys, authorization headers, endpoint credentials, file-system usernames, private audio, and personal text.

Do not commit real credentials or private recordings to this repository. Use placeholders in tests and examples.

## Third-party components

VoiceIME can download and use third-party runtimes, models, and cloud services. Vulnerabilities that originate entirely in an upstream component should also be reported to that component's maintainer. Reports about insecure integration, download, configuration, or use of those components in VoiceIME remain in scope here.

# Voice Speakers — Aliens

Chrome / Edge Manifest V3 extension for accurately extracting and speaking selected team-chat messages.

## What it does

- Accepts text the user explicitly copies/pastes from Microsoft Teams or another chat.
- Preserves message authors and message boundaries.
- Uses an optional OpenAI-compatible Chat Completions API for deeper attribution and faithful transcription.
- Falls back to local extraction when no API key is configured.
- Redacts common password, OTP, 2FA, access-token, AWS-key and API-key patterns before remote analysis.
- Speaks the final text using the browser's built-in SpeechSynthesis API.
- Provides popup, side panel and settings pages.

## Install locally

1. Open `chrome://extensions` (or `edge://extensions`).
2. Enable **Developer mode**.
3. Select **Load unpacked** and choose this `Voice-Speakers` directory.
4. Open the extension popup, paste copied chat text, and select **Analyze messages**.

## AI configuration

Open extension Settings and enter an API key you control. The extension uses the OpenAI Chat Completions endpoint with the configured model (default: `gpt-4.1-mini`). API usage may incur charges. The key is kept in `chrome.storage.local`, which is not encrypted by this project. Do not use a shared/team-wide key or paste secrets into chat input.

## Important limitations

- The extension does **not** silently read private Teams messages, access Teams databases, bypass authentication, or scrape hidden chat content. The user must explicitly copy/paste text.
- Local parsing is heuristic. For screenshots, quoted replies, nested threads, complex layouts, and ambiguous authors, AI analysis or manual verification is required.
- Browser speech depends on installed system voices and browser support. Speech synthesis is not a guaranteed offline service on every platform.
- No automatic recording, transcription of microphone audio, or background surveillance is included.

## Security

Treat pasted chat content as confidential. Review the text before submitting it to any AI provider. Keep redaction enabled. Redaction uses patterns and cannot guarantee detection of every secret format. Rotate any credentials exposed in chat screenshots or logs.

## Development structure

- `manifest.json` — extension definition
- `popup.*` — toolbar popup
- `sidepanel.html` / `ui.js` — side panel UI
- `service.js` — analysis orchestration and API call
- `analyzer.js` — local parser, redaction and system prompt
- `options.*` — API and speech settings

## Current status

Initial MVP scaffold. Load and smoke-test in Chrome/Edge before production use. The browser extension does not claim verified sender/reply threading for arbitrary Teams DOM or screenshots; it only reports relationships that the AI can ground in supplied text.
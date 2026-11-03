# Voice Speakers — Aliens

Chrome Manifest V3 extension that reads ChatGPT replies aloud.

## Current workflow

### Microsoft Teams live reader
1. Click the Voice Speakers toolbar icon. A compact popup opens; it does not open the side panel automatically.
2. Switch **Live message reading** to **ON** once.
3. The extension monitors all currently open Teams tabs at `teams.cloud.microsoft` and `teams.microsoft.com`, including tabs running in the background.
4. Switching to another browser tab does not stop the monitor. New Teams tabs are monitored automatically while the setting remains ON.
5. New incoming messages are spoken as **“Sender name ne kaha: message text.”**
6. Switch to **OFF** to stop monitoring and stop queued speech.

Only messages in Teams chat panes that are open/rendered in the browser are monitored. Hidden chats and unopened account-wide chat history are not fetched. When a Teams chat is changed, its already-loaded messages are used as a baseline and skipped.

## Install locally
1. Download or clone this repository.
2. Open chrome://extensions and enable Developer mode.
3. Choose **Load unpacked** and select the `Voice-Speakers` folder.
4. Open ChatGPT and click the extension icon.

## Notes
- Nested UI copies inside the same ChatGPT message are grouped as one prompt; separate prompts with identical wording remain separate when they are separate message turns.
- Reply pairing follows conversation message order: each prompt is linked to the assistant message(s) before the next user prompt.
- Extraction depends on ChatGPT's page structure and may need updates if its UI changes.
- Voice names and gender availability depend on the browser and operating system.
- The extension reads visible/rendered conversation messages only.
- No API key is needed for conversation loading or speech playback.
- Existing optional AI analysis settings remain available in the extension settings page.



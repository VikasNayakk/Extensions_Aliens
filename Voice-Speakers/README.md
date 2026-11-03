# Voice Speakers — Aliens

Chrome Manifest V3 extension that reads ChatGPT replies aloud.

Current Teams live reader release: **1.6.1**.

## Current workflow

### Microsoft Teams group reader
1. Open a Teams group in its own browser tab and click the Voice Speakers toolbar icon.
2. Click **Add currently open Teams group**. The group name is detected from the Teams tab.
3. When asked, choose **Yes, read last 5** or **No, only new**.
4. Add other selected groups from their normal Teams chats.
5. Move groups up or down to set priority.
6. Turn **Live message reading** ON.
7. When several messages from one group arrive together, the group name is announced once, followed by a five-second pause and then each message in order at a steady pace.
8. Repeated message/card IDs are stored and ignored. Emoji and standalone date/time lines are removed. Reply quote blocks are excluded when Teams exposes them; a reply is spoken as the original sender's name followed by the replier and the reply text.
9. Audio elements and direct audio-file links found inside a visible Teams message are played; speech waits until playback ends. The extension does not narrate the audio itself.
10. Turn OFF to stop monitoring, speech, and extension-started audio playback.

This browser-only mode watches Teams in-app notification cards in an open Teams tab and matches each notification's group name against your selected groups. It does not use Microsoft Graph, Entra app registration, Power Automate, or separate hidden monitor tabs. Keep at least one Teams web tab open. In Teams, enable message previews in notifications and make sure chat notifications are shown as banners. Existing notification cards are ignored when monitoring first starts; new cards are processed after that. If Teams does not create an in-app notification card for a chat, the extension cannot read that message in this mode. The one-time last-five option reads from the currently open chat when you add it.

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



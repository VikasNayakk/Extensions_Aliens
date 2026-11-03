# Voice Speakers — Aliens

Chrome Manifest V3 extension that reads ChatGPT replies aloud.

## Current workflow

### Microsoft Teams group reader
1. Open a Teams group in its own browser tab and click the Voice Speakers toolbar icon.
2. Click **Add currently open Teams group**. The group name is detected from the Teams tab.
3. When asked, choose **Yes, read last 5** or **No, only new**.
4. Add other groups from their own Teams tabs. Keep one Teams tab open per selected group so each group's live message list stays rendered.
5. Move groups up or down to set priority. The group at the top gets first place when multiple messages are waiting. A short one-second collection window lets messages arriving together be sorted before speech starts.
6. Turn **Live message reading** ON. Existing messages are skipped unless you explicitly choose to read the last five.
7. A live message is announced as: group name, then a five-second pause, then sender name and message. Messages are queued one by one by group priority.
8. Turn OFF to stop monitoring and clear queued speech.

The Teams web client only renders the chat currently open in each tab. Therefore each selected group needs its own Teams tab for simultaneous monitoring. Switching to non-Teams browser tabs does not stop monitoring. If the same selected group is opened in another Teams tab, its group name is matched to the selected-group list; tab identity is not used to block the message. A group chat navigation in its tab is treated as a new baseline to avoid reading old history as new.


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



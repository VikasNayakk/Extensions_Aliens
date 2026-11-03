# Voice Speakers — Aliens

Chrome Manifest V3 extension that reads ChatGPT replies aloud.

Current Teams live reader release: **1.5.7**.

## Current workflow

### Microsoft Teams group reader
1. Open a Teams group in its own browser tab and click the Voice Speakers toolbar icon.
2. Click **Add currently open Teams group**. The group name is detected from the Teams tab.
3. When asked, choose **Yes, read last 5** or **No, only new**.
4. Add other groups from their normal Teams chats. The extension creates a dedicated inactive/background Teams tab for each selected group, so you do not need to keep those chats open in your working tab.
5. Move groups up or down to set priority. The group at the top gets first place when multiple messages are waiting. A short one-second collection window lets messages arriving together be sorted before speech starts.
6. Turn **Live message reading** ON. Existing messages are skipped unless you explicitly choose to read the last five.
7. A live message is announced as: group name, then a five-second pause, then sender name and message. Messages are queued one by one by group priority.
8. Turn OFF to stop monitoring and clear queued speech.

The Teams web client renders the open chat in each tab. Voice Speakers now creates one dedicated inactive monitor tab per selected group using the chat ID and Teams deep link. Your normal Teams tab can open any other group; selected group monitors continue in their own background tabs. If a monitor tab is closed, the extension recreates it while ON. Existing groups saved before v1.5.8 do not yet have a chat ID/deep link, so re-add each selected group once to provision its dedicated background monitor.


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



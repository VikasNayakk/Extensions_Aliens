# Voice Speakers — Aliens

Chrome Manifest V3 extension that reads ChatGPT replies aloud.

## Current workflow
1. Open a ChatGPT conversation at chatgpt.com or chat.openai.com.
2. Click the Voice Speakers toolbar icon. The side panel opens and loads the active conversation.
3. The panel lists your prompts. The latest prompt is selected automatically.
4. Select any prompt or search your prompt text; its associated ChatGPT reply appears.
5. Press **Read selected reply** to hear the reply paired with that prompt, or **Read latest** for the newest prompt. Long replies are read in sequential chunks. Adjust speech rate and choose a browser-provided voice (default, female-sounding, or male-sounding).

Use **Refresh** after new messages arrive. The extension reads the currently open conversation's rendered page; it does not fetch your account-wide ChatGPT history or hidden conversations.

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


## Microsoft Teams live reading
1. Open Microsoft Teams in Chrome at `teams.cloud.microsoft` or `teams.microsoft.com`.
2. Open the chat you want to monitor and click the Voice Speakers toolbar icon.
3. In **Microsoft Teams — Live Reader**, click **Start live reading**.
4. Keep the side panel open. When a new incoming message appears in the currently open chat, the extension announces: **“Sender name ne kaha: message text.”**
5. Click **Stop** to stop monitoring and speech.

The monitor starts with a baseline of the messages already on screen, so it does not read old chat history. When you switch to another chat, the new chat's existing messages are skipped and monitoring continues for messages that arrive afterward. Messages sent by you are filtered where Teams exposes an outgoing-message marker. Common password, OTP, API-key and access-token patterns are masked before speech. The extension reads the open chat's rendered UI only; it does not read every chat in the Teams account or access hidden conversations.

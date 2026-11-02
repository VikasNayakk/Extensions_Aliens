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

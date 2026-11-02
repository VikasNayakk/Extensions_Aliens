const SECRET_PATTERNS = [
  /\b(?:password|passwd|pass)\s*[:=\-]\s*[^\s,;]+/gi,
  /\b(?:otp|2fa|totp|recovery\s*code|secret\s*key|access\s*token|api\s*key)\s*[:=\-]\s*[^\s,;]+/gi,
  /\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/g,
  /\bsk-[A-Za-z0-9_-]{16,}\b/g,
  /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi
];

export function redactSensitiveText(text) {
  let safe = String(text ?? "");
  for (const pattern of SECRET_PATTERNS) safe = safe.replace(pattern, "[REDACTED]");
  return safe;
}

const clean = s => String(s ?? "").replace(/\s+/g, " ").trim();
const knownName = s => clean(s).replace(/\s+\d{1,2}:\d{2}(?:\s*[AP]M)?(?:\s+Edited)?$/i, "").trim();

export function parseMessages(input) {
  const lines = String(input ?? "").replace(/\r/g, "").split("\n");
  const messages = [];
  let current = null;

  const header = /^(.*?)\s{1,}(\d{1,2}:\d{2}\s*(?:AM|PM)?)(?:\s+(Edited))?$/i;
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    const match = line.match(header);
    if (match && match[1].length < 70) {
      if (current) messages.push(current);
      current = { sender: knownName(match[1]), time: match[2], text: [] };
    } else if (current) {
      current.text.push(line);
    } else {
      current = { sender: "Sender not identified", time: "", text: [line] };
    }
  }
  if (current) messages.push(current);
  return messages.map(m => ({...m, text: m.text.join("\n").trim()})).filter(m => m.text);
}

export function formatConversation(messages) {
  return messages.map((m, i) => {
    const body = redactSensitiveText(m.text);
    return `To: ${m.sender}\n${body}`;
  }).join("\n\n");
}

export function localAnalyze(input) {
  const messages = parseMessages(input);
  if (!messages.length) return "Message text not readable.";
  return formatConversation(messages);
}

export const SYSTEM_PROMPT = `You are a precise team-chat message transcriber and attribution assistant. Output ONLY message text in the exact compact format below. No introduction, summary, advice, analysis labels, or closing lines.
For each original message:
To: [Exact sender]
[Faithful message transcription]
For a direct reply, nest it under the original only when the quoted/reply relationship is explicit:
[Original sender]'s message — Reply by [Replying sender]
[Reply text]
Rules:
- Preserve sender names, message wording, IDs, times, quantities, task instructions, and status exactly.
- Never invent a recipient or reply relationship. If the recipient is not explicit, do not add "To" beyond the sender heading.
- Keep different messages separate. Do not turn a request into an action already completed.
- If text is unclear, write [unclear] only for that span.
- Do not interpret ambiguous screenshots or missing context as facts.
- Never reproduce credentials, passwords, OTPs, 2FA codes, API keys, access tokens, or private keys; replace them with [REDACTED].
- Output only the message text. Use the same language as the source.`;
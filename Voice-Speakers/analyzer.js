const SECRET_PATTERNS = [
  /\b(?:password|passwd|pass)\s*[:=\-]\s*[^\s,;]+/gi,
  /\b(?:otp|2fa|totp|recovery\s*code|secret\s*key|access\s*token|api\s*key)\s*[:=\-]\s*[^\s,;]+/gi,
  /\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/g,
  /\bsk-[A-Za-z0-9_-]{16,}\b/g
];
export function redactSensitiveText(text) {
  let safe=String(text??"");
  for(const p of SECRET_PATTERNS)safe=safe.replace(p,"[REDACTED]");
  return safe;
}
const HEADER=/^(.*?)\s{1,}(\d{1,2}:\d{2}\s*(?:AM|PM)?)(?:\s+(Edited))?$/i;
export function parseMessages(input) {
  const lines=String(input??"").replace(/\r/g,"").split("\n");
  const messages=[];let current=null;
  for(const raw of lines){
    const line=raw.trim();if(!line)continue;
    const match=line.match(HEADER);
    if(match&&match[1].trim().length>0&&match[1].length<70){
      if(current)messages.push(current);
      current={sender:match[1].trim(),time:match[2],text:[]};
    }else if(current){current.text.push(line);}
    else{current={sender:"Sender not identified",time:"",text:[line]};}
  }
  if(current)messages.push(current);
  return messages.map(m=>({...m,text:m.text.join("\n").trim()})).filter(m=>m.text);
}
export function formatConversation(messages){
  return messages.map(m=>`To: ${m.sender}\n${redactSensitiveText(m.text)}`).join("\n\n");
}
export function localAnalyze(input){
  const messages=parseMessages(input);
  return messages.length?formatConversation(messages):"Message text not readable.";
}
export const SYSTEM_PROMPT=`You transcribe and attribute team-chat messages. Output ONLY message text, no introduction, summary, advice, analysis or closing.
For each original message:
To: [Exact sender]
[Faithful message transcription]
For a direct reply, nest it under the original only when the quoted/reply relationship is explicit:
[Original sender]'s message — Reply by [Replying sender]
[Reply text]
Preserve wording, names, IDs, times, quantities and status. Never invent a recipient or reply link. Keep each message separate. If text is unclear, mark only that span [unclear]. Never reveal passwords, OTP/2FA codes, keys or tokens; use [REDACTED]. Use the source language. Output only message text.`;
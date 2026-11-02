import {localAnalyze,redactSensitiveText,SYSTEM_PROMPT} from "./analyzer.js";

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if(message?.type !== "ANALYZE_MESSAGES") return;
  analyze(message.text).then(sendResponse).catch(error=>sendResponse({output:"Analysis error: "+error.message,remote:false}));
  return true;
});

async function analyze(input) {
  const settings = await chrome.storage.local.get(["apiKey","model","redact"]);
  const safeText = settings.redact === false ? String(input) : redactSensitiveText(input);
  if(!settings.apiKey) return {output:localAnalyze(safeText),remote:false};
  const response = await fetch("https://api.openai.com/v1/chat/completions",{
    method:"POST",
    headers:{"Authorization":"Bearer "+settings.apiKey,"Content-Type":"application/json"},
    body:JSON.stringify({model:settings.model||"gpt-4.1-mini",temperature:0,messages:[{role:"system",content:SYSTEM_PROMPT},{role:"user",content:safeText}]})
  });
  if(!response.ok) throw new Error("AI API returned HTTP "+response.status);
  const data=await response.json();
  const output=data.choices?.[0]?.message?.content?.trim();
  if(!output) throw new Error("AI response was empty");
  return {output,remote:true};
}
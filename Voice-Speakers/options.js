const $=id=>document.getElementById(id);
async function load(){const v=await chrome.storage.local.get(["apiKey","model","voiceRate","redact"]);$("apiKey").value=v.apiKey||"";$("model").value=v.model||"gpt-4.1-mini";$("voiceRate").value=v.voiceRate??1;$("rateValue").textContent=Number($("voiceRate").value).toFixed(1)+"×";$("redact").checked=v.redact!==false;}
$("voiceRate").addEventListener("input",()=> $("rateValue").textContent=Number($("voiceRate").value).toFixed(1)+"×");
$("save").addEventListener("click",async()=>{await chrome.storage.local.set({apiKey:$("apiKey").value.trim(),model:$("model").value.trim()||"gpt-4.1-mini",voiceRate:Number($("voiceRate").value),redact:$("redact").checked});$("status").textContent="Settings saved.";});
$("clear").addEventListener("click",async()=>{$("apiKey").value="";await chrome.storage.local.remove("apiKey");$("status").textContent="API key cleared.";});
load();
const $ = id => document.getElementById(id);
$("readClipboard").addEventListener("click", async () => { try { $("chatText").value = await navigator.clipboard.readText(); $("status").textContent="Clipboard text pasted."; } catch { $("status").textContent="Use Ctrl+V to paste."; } });
$("analyze").addEventListener("click", async () => {
 const text=$("chatText").value.trim(); if(!text){$("status").textContent="Paste messages first.";return;}
 $("analyze").disabled=true;$("status").textContent="Analyzing…";
 try{const r=await chrome.runtime.sendMessage({type:"ANALYZE_MESSAGES",text});$("result").textContent=r.output;$("resultWrap").hidden=false;$("status").textContent=r.remote?"AI analysis complete.":"Local extraction complete.";}
 catch(e){$("status").textContent="Analysis failed: "+e.message;}finally{$("analyze").disabled=false;}
});
$("copy").addEventListener("click",()=>navigator.clipboard.writeText($("result").textContent));
$("speak").addEventListener("click",async()=>{speechSynthesis.cancel();const {voiceRate=1}=await chrome.storage.local.get("voiceRate");for(const c of ($("result").textContent.match(/[\s\S]{1,1800}/g)||[])){const u=new SpeechSynthesisUtterance(c);u.rate=voiceRate;speechSynthesis.speak(u);}});
$("stop").addEventListener("click",()=>speechSynthesis.cancel());
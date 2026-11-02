const $ = id => document.getElementById(id);
$("settings").addEventListener("click", () => chrome.runtime.openOptionsPage());
$("openPanel").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({active:true,currentWindow:true});
  if (tab?.id) chrome.sidePanel.open({tabId:tab.id});
});
$("readClipboard").addEventListener("click", async () => {
  try { $("chatText").value = await navigator.clipboard.readText(); $("status").textContent = "Clipboard text pasted."; }
  catch { $("status").textContent = "Clipboard access denied. Paste with Ctrl+V."; }
});
$("analyze").addEventListener("click", runAnalysis);
$("copy").addEventListener("click", () => navigator.clipboard.writeText($("result").textContent));
$("speak").addEventListener("click", speak);
$("stop").addEventListener("click", () => speechSynthesis.cancel());

async function runAnalysis() {
  const text = $("chatText").value.trim();
  if (!text) { $("status").textContent = "Paste messages first."; return; }
  $("analyze").disabled = true;
  $("status").textContent = "Analyzing…";
  try {
    const result = await chrome.runtime.sendMessage({type:"ANALYZE_MESSAGES", text});
    $("result").textContent = result.output;
    $("resultWrap").hidden = false;
    $("status").textContent = result.remote ? "AI analysis complete." : "Local structured extraction complete.";
  } catch (e) { $("status").textContent = "Analysis failed: " + e.message; }
  finally { $("analyze").disabled = false; }
}
async function speak() {
  const text = $("result").textContent;
  if (!text) return;
  speechSynthesis.cancel();
  const {voiceRate=1} = await chrome.storage.local.get("voiceRate");
  const chunks = text.match(/[\s\S]{1,1800}/g) || [];
  for (const chunk of chunks) {
    const utterance = new SpeechSynthesisUtterance(chunk);
    utterance.rate = voiceRate;
    speechSynthesis.speak(utterance);
  }
}
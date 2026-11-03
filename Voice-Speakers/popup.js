const $ = id => document.getElementById(id);

function paint(data) {
  const enabled=!!data.teamsLiveEnabled;
  applyTheme(data.teamsPopupTheme || "light");
  $("statusPill").textContent=enabled?"ON":"OFF";
  $("statusPill").className="pill "+(enabled?"on":"off");
  $("power").className="switch "+(enabled?"on":"off");
  $("power").setAttribute("aria-pressed",String(enabled));
  $("power").setAttribute("aria-label",enabled?"Turn Teams live reading off":"Turn Teams live reading on");
  $("power").disabled=false;
  $("tabCount").textContent=String(data.teamsTabCount||0);
  $("readCount").textContent=String(data.teamsReadCount||0);
  $("activityState").textContent=enabled?"Listening":"Paused";
  $("status").textContent=data.teamsLastStatus||(enabled?"Monitoring open Teams tabs, including background tabs.":"Live reading is OFF.");
  const events=$("events");
  events.replaceChildren();
  const list=Array.isArray(data.teamsRecentEvents)?data.teamsRecentEvents:[];
  if(!list.length){const empty=document.createElement("div");empty.className="empty";empty.textContent="No new messages yet.";events.append(empty);}
  else list.slice(0,5).forEach(entry=>{const row=document.createElement("div");row.className="event";const sender=document.createElement("strong");sender.textContent=entry.sender||"Unknown sender";const message=document.createElement("span");message.textContent=entry.text||"";row.append(sender,message);events.append(row);});
}

async function refresh(){try{const state=await chrome.runtime.sendMessage({type:"GET_TEAMS_LIVE_STATUS"});paint(state||{});}catch(e){$("status").textContent="Could not read extension state: "+e.message;}}
  
function applyTheme(theme) {
  const selected = theme === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = selected;
  const button = $("theme");
  if (button) {
    button.textContent = selected === "dark" ? "Light" : "Dark";
    button.setAttribute("aria-label", selected === "dark" ? "Switch to light mode" : "Switch to dark mode");
  }
}

$("theme").addEventListener("click", async () => {
  const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
  applyTheme(next);
  await chrome.storage.local.set({teamsPopupTheme:next});
});

$("power").addEventListener("click",async()=>{
  const next=$("power").getAttribute("aria-pressed")!=="true";
  $("power").disabled=true;
  $("status").textContent=next?"Turning live reading ON…":"Turning live reading OFF…";
  try{const state=await chrome.runtime.sendMessage({type:"SET_TEAMS_LIVE",enabled:next});paint(state||{});}
  catch(e){$("status").textContent="Unable to change state: "+e.message;}
  finally{$("power").disabled=false;}
});
$("settings").addEventListener("click",()=>chrome.runtime.openOptionsPage());
chrome.runtime.onMessage.addListener(message=>{if(message?.type==="TEAMS_LIVE_STATE_UPDATED")paint(message.state||{});});
chrome.storage.onChanged.addListener((changes,area)=>{if(area==="local" && Object.keys(changes).some(k=>["teamsLiveEnabled","teamsReadCount","teamsRecentEvents","teamsLastStatus","teamsPopupTheme"].includes(k)))refresh();});
refresh();
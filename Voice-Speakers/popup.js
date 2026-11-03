const $=id=>document.getElementById(id);
let popupState={};
let pendingGroupKey=null;

function applyTheme(theme){
  const selected=theme==="dark"?"dark":"light";
  document.documentElement.dataset.theme=selected;
  $("theme").textContent=selected==="dark"?"Light":"Dark";
  $("theme").setAttribute("aria-label",selected==="dark"?"Switch to light mode":"Switch to dark mode");
}

function paint(data){
  popupState=data||{};
  applyTheme(data.teamsPopupTheme||"light");
  const enabled=!!data.teamsLiveEnabled;
  $("statusPill").textContent=enabled?"ON":"OFF";
  $("statusPill").className="pill "+(enabled?"on":"off");
  $("power").className="switch "+(enabled?"on":"off");
  $("power").setAttribute("aria-pressed",String(enabled));
  $("power").setAttribute("aria-label",enabled?"Turn Teams live reading off":"Turn Teams live reading on");
  $("tabCount").textContent=String(data.teamsTabCount||0);
  $("readCount").textContent=String(data.teamsReadCount||0);
  $("queueCount").textContent=String(data.teamsQueueCount||0);
  $("groupCount").textContent=(data.teamsSelectedGroups||[]).length+" groups";
  $("activityState").textContent=enabled?"Listening":"Paused";
  if(data.teamsLastStatus) $("status").textContent=data.teamsLastStatus;

  renderGroups(data.teamsSelectedGroups||[]);
  renderEvents(data.teamsRecentEvents||[]);
}

function renderGroups(groups){
  const root=$("groups");root.replaceChildren();
  if(!groups.length){
    const empty=document.createElement("div");empty.className="empty";empty.textContent="No groups selected. Open a Teams group and add it here.";root.append(empty);return;
  }
  groups.forEach((group,index)=>{
    const row=document.createElement("div");row.className="group-row";
    const info=document.createElement("div");info.className="group-info";
    const title=document.createElement("strong");title.textContent=(index+1)+". "+group.name;
    const detail=document.createElement("small");detail.textContent=(group.tabIds||[]).length+" linked Teams tab(s)";
    info.append(title,detail);
    const controls=document.createElement("div");controls.className="group-controls";
    const up=document.createElement("button");up.type="button";up.textContent="↑";up.title="Higher priority";up.disabled=index===0;up.addEventListener("click",()=>moveGroup(index,-1));
    const down=document.createElement("button");down.type="button";down.textContent="↓";down.title="Lower priority";down.disabled=index===groups.length-1;down.addEventListener("click",()=>moveGroup(index,1));
    const remove=document.createElement("button");remove.type="button";remove.textContent="×";remove.title="Remove group";remove.className="remove";remove.addEventListener("click",()=>removeGroup(group.key));
    controls.append(up,down,remove);row.append(info,controls);root.append(row);
  });
}

function renderEvents(events){
  const root=$("events");root.replaceChildren();
  if(!events.length){const empty=document.createElement("div");empty.className="empty";empty.textContent="No new messages yet.";root.append(empty);return;}
  events.slice(0,5).forEach(entry=>{
    const row=document.createElement("div");row.className="event";
    const group=document.createElement("strong");group.textContent=entry.groupName||"Teams group";
    const sender=document.createElement("b");sender.textContent=entry.sender||"Unknown sender";
    const message=document.createElement("span");message.textContent=entry.text||"";
    row.append(group,sender,message);root.append(row);
  });
}

async function refresh(){
  try{const state=await chrome.runtime.sendMessage({type:"GET_TEAMS_LIVE_STATUS"});paint(state||{});}
  catch(error){$("status").textContent="Could not load status: "+error.message;}
}

$("power").addEventListener("click",async()=>{
  const next=$("power").getAttribute("aria-pressed")!=="true";
  $("power").disabled=true;
  try{const state=await chrome.runtime.sendMessage({type:"SET_TEAMS_LIVE",enabled:next});paint(state||{});}
  catch(error){$("status").textContent="Could not change status: "+error.message;}
  finally{$("power").disabled=false;}
});

$("addGroup").addEventListener("click",async()=>{
  $("addGroup").disabled=true;$("status").textContent="Detecting the currently open Teams group…";
  try{
    const result=await chrome.runtime.sendMessage({type:"ADD_CURRENT_TEAMS_GROUP"});
    if(!result?.ok) throw new Error(result?.error||"Could not add this Teams group.");
    pendingGroupKey=result.group.key;
    $("historyTitle").textContent='Read the last 5 messages from "'+result.group.name+'"?';
    $("historyPrompt").hidden=false;
    $("status").textContent='Added "'+result.group.name+'". Choose Yes or No for its history.';
    if(result.state) paint(result.state);
    $("historyPrompt").hidden=false;
  }catch(error){$("status").textContent=error.message;}
  finally{$("addGroup").disabled=false;}
});

$("historyYes").addEventListener("click",async()=>{
  if(!pendingGroupKey)return;
  $("historyYes").disabled=true;$("historyNo").disabled=true;
  $("status").textContent="Loading the last five messages…";
  try{
    const result=await chrome.runtime.sendMessage({type:"READ_LAST_FIVE_TEAMS_MESSAGES",groupKey:pendingGroupKey});
    if(result?.error)throw new Error(result.error);
    $("status").textContent="Last five messages queued in priority order.";
    if(result)paint(result);
  }catch(error){$("status").textContent=error.message;}
  finally{$("historyPrompt").hidden=true;$("historyYes").disabled=false;$("historyNo").disabled=false;pendingGroupKey=null;}
});

$("historyNo").addEventListener("click",async()=>{
  $("historyPrompt").hidden=true;
  $("status").textContent="Saved. Only new messages will be read from this group.";
  pendingGroupKey=null;
});

async function moveGroup(index,delta){
  const groups=[...(popupState.teamsSelectedGroups||[])];
  const target=index+delta;if(target<0||target>=groups.length)return;
  [groups[index],groups[target]]=[groups[target],groups[index]];
  const result=await chrome.runtime.sendMessage({type:"REORDER_TEAMS_GROUPS",order:groups.map(group=>group.key)});
  if(result)paint(result);
}
async function removeGroup(groupKey){
  const result=await chrome.runtime.sendMessage({type:"REMOVE_TEAMS_GROUP",groupKey});
  if(result)paint(result);
}

$("theme").addEventListener("click",async()=>{
  const next=document.documentElement.dataset.theme==="dark"?"light":"dark";
  applyTheme(next);await chrome.storage.local.set({teamsPopupTheme:next});
});
chrome.runtime.onMessage.addListener(message=>{
  if(message?.type==="TEAMS_LIVE_STATE_UPDATED")paint(message.state||{});
});
chrome.storage.onChanged.addListener((changes,area)=>{
  if(area==="local"&&Object.keys(changes).some(key=>[
    "teamsLiveEnabled","teamsReadCount","teamsRecentEvents","teamsLastStatus",
    "teamsPopupTheme","teamsSelectedGroups","teamsSpeechQueue"
  ].includes(key)))refresh();
});
refresh();
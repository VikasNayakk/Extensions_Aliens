import {localAnalyze,redactSensitiveText,SYSTEM_PROMPT} from "./analyzer.js";

chrome.runtime.onMessage.addListener((m,s,send)=>{if(m?.type==="ANALYZE_MESSAGES"){analyze(m.text).then(send).catch(e=>send({output:"Analysis error: "+e.message,remote:false}));return true;}if(m?.type==="TEST_OPENAI"){testApi().then(send).catch(e=>send({ok:false,error:e.message}));return true;}});
async function requestOpenAI(key,model,messages){let r;try{r=await fetch("https://api.openai.com/v1/chat/completions",{method:"POST",headers:{"Authorization":"Bearer "+key,"Content-Type":"application/json"},body:JSON.stringify({model,temperature:0,messages})});}catch(e){throw Error("Network/CORS error: "+e.message);}const data=await r.json().catch(()=>({}));if(!r.ok){const detail=data?.error?.message||data?.error?.code||("HTTP "+r.status);throw Error("OpenAI API "+r.status+": "+detail);}const output=data.choices?.[0]?.message?.content?.trim();if(!output)throw Error("OpenAI returned an empty response.");return output;}
async function testApi(){const s=await chrome.storage.local.get(["apiKey","model"]);if(!s.apiKey)throw Error("Save an API key first.");await requestOpenAI(s.apiKey,s.model||"gpt-4.1-mini",[{role:"user",content:"Reply with exactly: API connection successful"}]);return {ok:true,model:s.model||"gpt-4.1-mini"};}
async function analyze(input){const s=await chrome.storage.local.get(["apiKey","model","redact"]);const safe=s.redact===false?String(input):redactSensitiveText(input);if(!s.apiKey)return {output:localAnalyze(safe),remote:false};const output=await requestOpenAI(s.apiKey,s.model||"gpt-4.1-mini",[{role:"system",content:SYSTEM_PROMPT},{role:"user",content:safe}]);return {output,remote:true};}


const TEAMS_LIVE_KEY = "teamsLiveEnabled";
const TEAMS_GROUPS_KEY = "teamsSelectedGroups";
const TEAMS_READ_COUNT = "teamsReadCount";
const TEAMS_RECENT_EVENTS = "teamsRecentEvents";
const recentMessageIds = new Set();
let speechQueue = [];
let speechCurrent = null;
let speechQueueLoaded = false;
let speechProcessing = false;
let speechDispatchTimer = null;
let groupDelayTimer = null;

const teamsUrl = url => /^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(url || "");
const normalizeGroup = value => String(value || "").normalize("NFKC").replace(/\s+/g," ").trim().toLowerCase();

async function getTeamsTabs() {
  const tabs=await chrome.tabs.query({});
  return tabs.filter(tab=>teamsUrl(tab.url));
}

async function getStoredGroups() {
  const saved=await chrome.storage.local.get(TEAMS_GROUPS_KEY);
  return Array.isArray(saved[TEAMS_GROUPS_KEY])?saved[TEAMS_GROUPS_KEY]:[];
}

async function getLiveState() {
  const saved=await chrome.storage.local.get([
    TEAMS_LIVE_KEY,TEAMS_READ_COUNT,TEAMS_RECENT_EVENTS,
    "teamsLastStatus","teamsPopupTheme","teamsSpeechQueue"
  ]);
  const tabs=await getTeamsTabs();
  const groups=await getStoredGroups();
  return {
    teamsLiveEnabled:!!saved[TEAMS_LIVE_KEY],
    teamsTabCount:tabs.length,
    teamsGroupCount:groups.length,
    teamsSelectedGroups:groups,
    teamsReadCount:saved[TEAMS_READ_COUNT]||0,
    teamsRecentEvents:saved[TEAMS_RECENT_EVENTS]||[],
    teamsQueueCount:Array.isArray(saved.teamsSpeechQueue)?saved.teamsSpeechQueue.length:0,
    teamsLastStatus:saved.teamsLastStatus||"",
    teamsPopupTheme:saved.teamsPopupTheme||"light"
  };
}

async function publishLiveState() {
  const state=await getLiveState();
  chrome.runtime.sendMessage({type:"TEAMS_LIVE_STATE_UPDATED",state},()=>void chrome.runtime.lastError);
  return state;
}

async function setStatus(text) {
  await chrome.storage.local.set({teamsLastStatus:text});
  await publishLiveState();
}

async function injectTeamsTabs() {
  const tabs=await getTeamsTabs();
  let injected=0,failed=0;
  for(const tab of tabs) {
    try {
      await chrome.scripting.executeScript({target:{tabId:tab.id,allFrames:true},files:["teams-content.js"]});
      injected++;
    } catch (_) { failed++; }
  }
  return {total:tabs.length,injected,failed};
}

async function setTeamsLiveEnabled(enabled) {
  await chrome.storage.local.set({[TEAMS_LIVE_KEY]:!!enabled});
  if(!enabled) {
    if(speechDispatchTimer) clearTimeout(speechDispatchTimer);
    if(groupDelayTimer) clearTimeout(groupDelayTimer);
    speechDispatchTimer=null;
    groupDelayTimer=null;
    chrome.tts.stop();
    speechQueue=[];
    speechCurrent=null;
    speechProcessing=false;
    speechQueueLoaded=true;
    await chrome.storage.local.set({
      teamsSpeechQueue:[],
      teamsSpeechCurrent:null,
      teamsLastStatus:"Live reading is OFF. Speech queue cleared."
    });
  } else {
    const groups=await getStoredGroups();
    const result=await injectTeamsTabs();
    const status=groups.length
      ? "Live reading is ON. "+groups.length+" selected group(s); "+result.total+" Teams tab(s) open. Add one Teams tab per group."
      : "Live reading is ON, but no groups are selected. Add the current Teams group first.";
    await chrome.storage.local.set({teamsLastStatus:status});
  }
  return publishLiveState();
}

async function activeTeamsTab() {
  const [tab]=await chrome.tabs.query({active:true,lastFocusedWindow:true});
  if(!tab || !tab.id || !teamsUrl(tab.url)) throw new Error("Open the Teams group you want to add in the active browser tab first.");
  return tab;
}

async function addCurrentGroup() {
  const tab=await activeTeamsTab();
  let info;
  try {
    info=await chrome.tabs.sendMessage(tab.id,{type:"TEAMS_GET_CHAT_INFO"},{frameId:0});
  } catch (_) {
    await chrome.scripting.executeScript({target:{tabId:tab.id},files:["teams-content.js"]});
    info=await chrome.tabs.sendMessage(tab.id,{type:"TEAMS_GET_CHAT_INFO"},{frameId:0});
  }
  if(!info?.ok || !info.groupName || !info.groupKey) {
    throw new Error(info?.error || "Could not detect the current Teams group name. Open an individual group chat and retry.");
  }

  const groups=await getStoredGroups();
  let group=groups.find(item=>item.key===info.groupKey);
  if(group) {
    const tabIds=Array.isArray(group.tabIds)?group.tabIds:[];
    if(!tabIds.includes(tab.id)) tabIds.push(tab.id);
    group.tabIds=tabIds;
  } else {
    group={
      id:info.groupKey,
      key:info.groupKey,
      name:info.groupName,
      tabIds:[tab.id],
      priority:groups.length,
      readHistory:false
    };
    groups.push(group);
  }
  groups.forEach((item,index)=>item.priority=index);
  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  await setStatus("Added group: "+group.name+". Choose Yes or No to read its last five messages.");
  return {ok:true,group,baselineCount:info.baselineCount||0};
}

async function removeGroup(groupKey) {
  const groups=(await getStoredGroups()).filter(group=>group.key!==groupKey);
  groups.forEach((group,index)=>group.priority=index);
  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  await hydrateSpeechQueue();
  speechQueue=speechQueue.filter(item=>item.groupKey!==groupKey);
  await persistSpeechQueue();
  await setStatus("Group removed. Its queued messages were discarded.");
  return publishLiveState();
}

async function reorderGroups(order) {
  const groups=await getStoredGroups();
  const sorted=[];
  for(const key of order||[]) {
    const group=groups.find(item=>item.key===key);
    if(group && !sorted.includes(group)) sorted.push(group);
  }
  for(const group of groups) if(!sorted.includes(group)) sorted.push(group);
  sorted.forEach((group,index)=>group.priority=index);
  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:sorted});
  await setStatus("Group priority order saved.");
  return publishLiveState();
}

async function readLastFive(groupKey) {
  const groups=await getStoredGroups();
  const group=groups.find(item=>item.key===groupKey);
  if(!group) throw new Error("This group is no longer selected.");
  const tabId=(group.tabIds||[]).find(id=>Number.isInteger(id));
  if(!tabId) throw new Error("No Teams tab is linked to this group. Re-add the group from its Teams tab.");
  const allTabs=await chrome.tabs.query({});
  if(!allTabs.some(item=>item.id===tabId && teamsUrl(item.url))) throw new Error("The Teams tab for this group is closed. Open it again and re-add the group.");
  let result;
  try {
    result=await chrome.tabs.sendMessage(tabId,{type:"TEAMS_READ_LAST_FIVE",groupKey},{frameId:0});
  } catch (_) {
    await chrome.scripting.executeScript({target:{tabId},files:["teams-content.js"]});
    result=await chrome.tabs.sendMessage(tabId,{type:"TEAMS_READ_LAST_FIVE",groupKey},{frameId:0});
  }
  if(!result?.ok) throw new Error(result?.error||"Could not read the last five messages from this group.");
  const messages=(result.messages||[]).slice(-5);
  if(!messages.length) {
    await setStatus("No readable incoming messages were found in "+group.name+".");
    return publishLiveState();
  }
  const items=[{
    groupKey:group.key,
    groupName:group.name,
    messages:messages.map(message=>({sender:message.sender,text:message.text,messageId:message.key})),
    historyBatch:true,
    priority:groups.indexOf(group),
    announceGroup:true,
    sequence:Date.now()
  }];
  await enqueueSpeech(items);
  group.readHistory=true;
  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  await setStatus("Queued "+messages.length+" recent message(s) from "+group.name+".");
  return publishLiveState();
}

async function hydrateSpeechQueue() {
  if(speechQueueLoaded) return;
  const saved=await chrome.storage.local.get(["teamsSpeechQueue","teamsSpeechCurrent"]);
  speechQueue=Array.isArray(saved.teamsSpeechQueue)?saved.teamsSpeechQueue:[];
  speechCurrent=saved.teamsSpeechCurrent||null;
  speechQueueLoaded=true;
  if(speechCurrent?.phase==="gap") {
    const sequence=speechCurrent.sequence;
    groupDelayTimer=setTimeout(()=>{
      if(speechCurrent?.sequence===sequence && speechCurrent.phase==="gap") {
        speechCurrent.phase="message";
        persistSpeechQueue().then(speakCurrentPhase);
      }
    },5000);
  }
}

async function persistSpeechQueue() {
  await chrome.storage.local.set({
    teamsSpeechQueue:speechQueue,
    teamsSpeechCurrent:speechCurrent
  });
}

async function enqueueSpeech(items) {
  await hydrateSpeechQueue();
  speechQueue.push(...items);
  await persistSpeechQueue();
  scheduleSpeechQueue();
}

function scheduleSpeechQueue() {
  if(speechCurrent || speechDispatchTimer || !speechQueue.length) return;
  speechDispatchTimer=setTimeout(()=>{
    speechDispatchTimer=null;
    processSpeechQueue();
  },1000);
}

async function currentGroupPriority(groupKey) {
  const groups=await getStoredGroups();
  const index=groups.findIndex(group=>group.key===groupKey);
  return index<0?999:index;
}

async function processSpeechQueue() {
  if(speechProcessing) return;
  await hydrateSpeechQueue();
  if(speechCurrent || !speechQueue.length) return;
  speechProcessing=true;
  try {
    for(const item of speechQueue) item.priority=await currentGroupPriority(item.groupKey);
    speechQueue.sort((a,b)=>(a.priority-b.priority)||(a.sequence-b.sequence));
    speechCurrent=speechQueue.shift();
    speechCurrent.phase=speechCurrent.announceGroup?"group":"message";
    await persistSpeechQueue();
    speakCurrentPhase();
  } finally {
    speechProcessing=false;
  }
}

function speakCurrentPhase() {
  if(!speechCurrent) return;
  const current=speechCurrent;
  const historyMessage=current.historyBatch
    ? current.messages[Math.min(current.messageIndex||0,current.messages.length-1)]
    : null;
  const text=current.phase==="group"
    ? current.groupName
    : historyMessage
      ? historyMessage.sender+" ne kaha: "+historyMessage.text
      : current.sender+" ne kaha: "+current.text;

  const speakWithVoice = availableVoices => {
    if(!speechCurrent || speechCurrent.sequence!==current.sequence) return;
    const voices=Array.isArray(availableVoices)?availableVoices:[];
    const voicesWithEnd=voices.filter(item=>Array.isArray(item.eventTypes)&&item.eventTypes.includes("end"));
    const voice=
      voicesWithEnd.find(item=>/^hi[-_]IN$/i.test(item.lang||"")) ||
      voicesWithEnd.find(item=>/^hi(?:[-_]|$)/i.test(item.lang||"")) ||
      voicesWithEnd.find(item=>/^en[-_]IN$/i.test(item.lang||"")) ||
      null;
    const options={
      enqueue:false,
      rate:0.85,
      requiredEventTypes:["end"],
      desiredEventTypes:["start","end","error","interrupted","cancelled"],
      onEvent:event=>{
        if(!speechCurrent || speechCurrent.sequence!==current.sequence) return;
        if(event.type==="start") {
          setStatus("Speaking "+(current.groupName||"Teams group")+" · "+(current.phase==="group"?"group name":(historyMessage?.sender||current.sender)));
        } else if(event.type==="end") {
          if(current.phase==="group") {
            current.phase="gap";
            persistSpeechQueue().then(()=>{
              if(groupDelayTimer) clearTimeout(groupDelayTimer);
              const sequence=current.sequence;
              groupDelayTimer=setTimeout(()=>{
                groupDelayTimer=null;
                if(speechCurrent?.sequence===sequence && speechCurrent.phase==="gap") {
                  speechCurrent.phase="message";
                  persistSpeechQueue().then(speakCurrentPhase);
                }
              },5000);
            });
            setStatus("Group "+current.groupName+" announced. Waiting 5 seconds before the message.");
          } else if(current.historyBatch && (current.messageIndex||0)+1<current.messages.length) {
            current.messageIndex=(current.messageIndex||0)+1;
            persistSpeechQueue().then(speakCurrentPhase);
          } else {
            finishCurrentSpeech();
          }
        } else if(["error","interrupted","cancelled"].includes(event.type)) {
          if(event.type==="error") setStatus("Speech error: "+(event.errorMessage||"Chrome TTS could not read the message."));
          finishCurrentSpeech();
        }
      }
    };
    if(voice?.voiceName) {
      options.voiceName=voice.voiceName;
      options.lang=voice.lang;
    } else {
      // Leave the voice unset if Hindi/India voices are not installed;
      // Chrome then falls back to its default available speech engine.
      setStatus("No Hindi voice is installed. Trying the default Chrome speech voice.");
    }
    chrome.tts.speak(text,options,()=>{
      if(chrome.runtime.lastError) {
        setStatus("Chrome TTS error: "+chrome.runtime.lastError.message);
        finishCurrentSpeech();
      }
    });
  };

  try {
    chrome.tts.getVoices(voices=>speakWithVoice(voices));
  } catch(error) {
    setStatus("Could not load speech voices: "+error.message);
    speakWithVoice([]);
  }
}

async function finishCurrentSpeech() {
  speechCurrent=null;
  if(groupDelayTimer) clearTimeout(groupDelayTimer);
  groupDelayTimer=null;
  await persistSpeechQueue();
  scheduleSpeechQueue();
}

async function onTeamsMessage(message,sender) {
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(!enabled[TEAMS_LIVE_KEY]) return;
  const groupKey=String(message.groupKey||"");
  const groups=await getStoredGroups();
  const group=groups.find(item=>item.key===groupKey);
  if(!group) {
    await setStatus("Ignored Teams message: this chat is not in Selected groups (received group key: "+groupKey+").");
    return;
  }
  const tabId=sender?.tab?.id;
  let sourceUrl=sender?.url || sender?.tab?.url || "";
  if(!sourceUrl && Number.isInteger(tabId)) {
    try { sourceUrl=(await chrome.tabs.get(tabId)).url || ""; } catch (_) {}
  }
  if(!Number.isInteger(tabId) || !teamsUrl(sourceUrl)) {
    await setStatus("Message received from an unknown tab; cannot safely route it to speech.");
    return;
  }

  // Match selected groups by normalized chat identity, not the tab ID used when they were added.
  // This allows the same selected group to be read when it is open in another Teams tab.
  const senderName=String(message.sender||"").trim();
  const text=String(message.text||"").trim();
  const messageId=String(message.messageId||"");
  if(!senderName || !text) {
    await setStatus("Ignored Teams message: sender or message text was empty.");
    return;
  }
  const uniqueId=tabId+"|"+groupKey+"|"+messageId;
  if(messageId && recentMessageIds.has(uniqueId)) return;
  if(messageId) {
    recentMessageIds.add(uniqueId);
    if(recentMessageIds.size>5000) recentMessageIds.clear();
  }

  const entry={groupName:group.name,groupKey,sender:senderName,text,receivedAt:Date.now()};
  const saved=await chrome.storage.local.get([TEAMS_READ_COUNT,TEAMS_RECENT_EVENTS]);
  const recent=[entry,...(saved[TEAMS_RECENT_EVENTS]||[])].slice(0,8);
  await chrome.storage.local.set({
    [TEAMS_READ_COUNT]:(saved[TEAMS_READ_COUNT]||0)+1,
    [TEAMS_RECENT_EVENTS]:recent,
    teamsLastStatus:"New message in "+group.name+" from "+senderName+"."
  });
  await setStatus("Detected a new message in selected group: "+group.name+" · "+senderName+". Queued for priority speech.");
  await enqueueSpeech([{
    ...entry,
    priority:groups.indexOf(group),
    announceGroup:true,
    sequence:Date.now(),
    messageId
  }]);
  await publishLiveState();
}

chrome.runtime.onMessage.addListener((message,sender,sendResponse)=>{
  if(message?.type==="GET_TEAMS_LIVE_STATUS") {
    getLiveState().then(sendResponse).catch(error=>sendResponse({teamsLastStatus:error.message}));
    return true;
  }
  if(message?.type==="SET_TEAMS_LIVE") {
    setTeamsLiveEnabled(!!message.enabled).then(sendResponse).catch(error=>sendResponse({teamsLastStatus:error.message}));
    return true;
  }
  if(message?.type==="ADD_CURRENT_TEAMS_GROUP") {
    addCurrentGroup().then(async result=>sendResponse({...result,state:await publishLiveState()}))
      .catch(error=>sendResponse({ok:false,error:error.message}));
    return true;
  }
  if(message?.type==="REMOVE_TEAMS_GROUP") {
    removeGroup(message.groupKey).then(sendResponse).catch(error=>sendResponse({error:error.message}));
    return true;
  }
  if(message?.type==="REORDER_TEAMS_GROUPS") {
    reorderGroups(message.order).then(sendResponse).catch(error=>sendResponse({error:error.message}));
    return true;
  }
  if(message?.type==="READ_LAST_FIVE_TEAMS_MESSAGES") {
    readLastFive(message.groupKey).then(sendResponse).catch(error=>sendResponse({error:error.message}));
    return true;
  }
  if(message?.type==="TEAMS_LIVE_MESSAGE") onTeamsMessage(message,sender).catch(error=>setStatus("Teams message pipeline error: "+error.message));
  if(message?.type==="TEAMS_LIVE_STATUS") {
    const text=String(message.text||"");
    chrome.storage.local.set({teamsLastStatus:text});
  }
});

chrome.tabs.onUpdated.addListener(async(tabId,changeInfo,tab)=>{
  if(changeInfo.status!=="complete" || !teamsUrl(tab?.url)) return;
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY]) {
    try { await chrome.scripting.executeScript({target:{tabId,allFrames:true},files:["teams-content.js"]}); }
    catch (_) {}
  }
  publishLiveState();
});
chrome.tabs.onCreated.addListener(async tab=>{
  if(!teamsUrl(tab.url)) return;
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY] && Number.isInteger(tab.id)) {
    try { await chrome.scripting.executeScript({target:{tabId:tab.id,allFrames:true},files:["teams-content.js"]}); }
    catch (_) {}
  }
});
chrome.tabs.onRemoved.addListener(async tabId=>{
  const groups=await getStoredGroups();
  let changed=false;
  for(const group of groups) {
    const before=(group.tabIds||[]).length;
    group.tabIds=(group.tabIds||[]).filter(id=>id!==tabId);
    if(before!==group.tabIds.length) changed=true;
  }
  if(changed) await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  publishLiveState();
});
chrome.runtime.onStartup.addListener(async()=>{
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY]) await injectTeamsTabs();
});
chrome.runtime.onInstalled.addListener(async()=>{
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY]) await injectTeamsTabs();
});

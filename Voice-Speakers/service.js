import {localAnalyze,redactSensitiveText,SYSTEM_PROMPT} from "./analyzer.js";

chrome.runtime.onMessage.addListener((m,s,send)=>{if(m?.type==="ANALYZE_MESSAGES"){analyze(m.text).then(send).catch(e=>send({output:"Analysis error: "+e.message,remote:false}));return true;}if(m?.type==="TEST_OPENAI"){testApi().then(send).catch(e=>send({ok:false,error:e.message}));return true;}});
async function requestOpenAI(key,model,messages){let r;try{r=await fetch("https://api.openai.com/v1/chat/completions",{method:"POST",headers:{"Authorization":"Bearer "+key,"Content-Type":"application/json"},body:JSON.stringify({model,temperature:0,messages})});}catch(e){throw Error("Network/CORS error: "+e.message);}const data=await r.json().catch(()=>({}));if(!r.ok){const detail=data?.error?.message||data?.error?.code||("HTTP "+r.status);throw Error("OpenAI API "+r.status+": "+detail);}const output=data.choices?.[0]?.message?.content?.trim();if(!output)throw Error("OpenAI returned an empty response.");return output;}
async function testApi(){const s=await chrome.storage.local.get(["apiKey","model"]);if(!s.apiKey)throw Error("Save an API key first.");await requestOpenAI(s.apiKey,s.model||"gpt-4.1-mini",[{role:"user",content:"Reply with exactly: API connection successful"}]);return {ok:true,model:s.model||"gpt-4.1-mini"};}
async function analyze(input){const s=await chrome.storage.local.get(["apiKey","model","redact"]);const safe=s.redact===false?String(input):redactSensitiveText(input);if(!s.apiKey)return {output:localAnalyze(safe),remote:false};const output=await requestOpenAI(s.apiKey,s.model||"gpt-4.1-mini",[{role:"system",content:SYSTEM_PROMPT},{role:"user",content:safe}]);return {output,remote:true};}


const TEAMS_LIVE_KEY = "teamsLiveEnabled";
const TEAMS_GROUPS_KEY = "teamsSelectedGroups";
const TEAMS_READ_COUNT = "teamsReadCount";
const TEAMS_RECENT_EVENTS = "teamsRecentEvents";
const recentMessageIds = new Map();
let speechQueue = [];
let speechCurrent = null;
let speechQueueLoaded = false;
let speechProcessing = false;
let speechDispatchTimer = null;
let groupDelayTimer = null;
let speechAdvanceTimer = null;

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
    teamsMonitorCount:tabs.length,
    teamsMonitorMissing:[],
    teamsGroupCount:groups.length,
    teamsSelectedGroups:groups,
    teamsReadCount:saved[TEAMS_READ_COUNT]||0,
    teamsRecentEvents:saved[TEAMS_RECENT_EVENTS]||[],
    teamsQueueCount:Array.isArray(saved.teamsSpeechQueue)?saved.teamsSpeechQueue.reduce((total,batch)=>total+(Array.isArray(batch.messages)?batch.messages.length:(batch.sender?1:0)),0):0,
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
    if(speechAdvanceTimer) clearTimeout(speechAdvanceTimer);
    speechDispatchTimer=null;
    speechAdvanceTimer=null;
    groupDelayTimer=null;
    chrome.tts.stop();
    await stopTeamsAudio();
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
    for(const group of groups) {
      if(Number.isInteger(group.monitorTabId)) {
        try { await chrome.tabs.remove(group.monitorTabId); } catch (_) {}
        group.monitorTabId=null;
      }
      delete group.monitorUrl;
      delete group.chatId;
    }
    await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
    const result=await injectTeamsTabs();
    const status=result.total===0
      ? "Live reading is ON, but Teams is not open. Open Teams in one browser tab to monitor its in-app notifications."
      : "Live reading is ON. Watching Teams in-app notifications across "+result.injected+" Teams tab(s). Keep Teams open and allow message previews.";
    await chrome.storage.local.set({teamsLastStatus:status});
  }
  return publishLiveState();
}

async function activeTeamsTab() {
  const [tab]=await chrome.tabs.query({active:true,lastFocusedWindow:true});
  if(!tab || !tab.id || !teamsUrl(tab.url)) throw new Error("Open the Teams group you want to add in the active browser tab first.");
  return tab;
}

async function waitForMonitorTab(tabId, timeoutMs=25000) {
  const deadline=Date.now()+timeoutMs;
  while(Date.now()<deadline) {
    let tab=null;
    try { tab=await chrome.tabs.get(tabId); } catch (_) { return null; }
    if(tab?.discarded) {
      try { await chrome.tabs.reload(tabId); } catch (_) {}
      await new Promise(resolve=>setTimeout(resolve,350));
      continue;
    }
    if(tab && teamsUrl(tab.url)) {
      try { await chrome.tabs.update(tabId,{autoDiscardable:false}); } catch (_) {}
    }
    if(tab && teamsUrl(tab.url) && tab.status==="complete") return tab;
    await new Promise(resolve=>setTimeout(resolve,350));
  }
  throw new Error("The dedicated Teams monitor tab is still loading. Try again in a few seconds.");
}

async function createOrGetMonitorTab(group) {
  if(!group?.monitorUrl || !group?.chatId) {
    throw new Error("This saved group has no independent monitor link. Open this group and add it again once.");
  }

  if(Number.isInteger(group.monitorTabId)) {
    try {
      const existing=await chrome.tabs.get(group.monitorTabId);
      if(existing && teamsUrl(existing.url)) {
        await chrome.tabs.update(existing.id,{autoDiscardable:false});
        if(existing.discarded) await chrome.tabs.reload(existing.id);
        return existing.id;
      }
    } catch (_) {}
  }

  const created=await chrome.tabs.create({url:group.monitorUrl,active:false});
  group.monitorTabId=created.id;
  await chrome.tabs.update(created.id,{autoDiscardable:false});
  return created.id;
}

async function ensureGroupMonitorTabs() {
  const groups=await getStoredGroups();

  // Upgrade old saved groups automatically when their original Teams tab is still on that group.
  for(const group of groups) {
    if(group.chatId && group.monitorUrl) continue;
    for(const oldTabId of group.tabIds||[]) {
      try {
        const oldTab=await chrome.tabs.get(oldTabId);
        if(!teamsUrl(oldTab.url)) continue;
        await chrome.scripting.executeScript({
          target:{tabId:oldTabId,frameIds:[0]},
          files:["teams-content.js"]
        });
        const info=await chrome.tabs.sendMessage(oldTabId,{type:"TEAMS_GET_CHAT_INFO_V16"},{frameId:0});
        if(info?.ok && normalizeGroup(info.groupName)===normalizeGroup(group.name) && info.chatId && info.monitorUrl) {
          group.chatId=info.chatId;
          group.monitorUrl=info.monitorUrl;
          group.name=info.groupName;
          break;
        }
      } catch (_) {}
    }
  }
  const missing=[];
  const failed=[];
  let created=0;
  let ready=0;

  for(const group of groups) {
    if(!group.chatId || !group.monitorUrl) {
      missing.push(group.name);
      continue;
    }
    try {
      const previous=group.monitorTabId;
      const tabId=await createOrGetMonitorTab(group);
      if(tabId!==previous) created++;
      ready++;
    } catch(error) {
      failed.push(group.name+": "+error.message);
    }
  }

  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  for(const group of groups) {
    if(!Number.isInteger(group.monitorTabId)) continue;
    try {
      const tab=await chrome.tabs.get(group.monitorTabId);
      if(tab?.status==="complete" && teamsUrl(tab.url) && !tab.discarded) {
        await restoreMonitorTabToSelectedGroup(group,group.monitorTabId);
      }
    } catch (_) {}
  }

  return {total:groups.length,ready,created,missing,failed};
}

async function restoreMonitorTabToSelectedGroup(group, tabId) {
  if(!group?.monitorUrl || !group?.chatId || !Number.isInteger(tabId)) return false;
  try {
    await chrome.scripting.executeScript({
      target:{tabId,frameIds:[0]},
      files:["teams-content.js"]
    });
    const info=await chrome.tabs.sendMessage(tabId,{type:"TEAMS_GET_CHAT_INFO_V16"},{frameId:0});
    if(info?.ok && info.chatId && info.chatId!==group.chatId) {
      await chrome.tabs.update(tabId,{
        url:group.monitorUrl,
        active:false,
        autoDiscardable:false
      });
      return true;
    }
  } catch (_) {}
  return false;
}

async function sendToMonitorTab(tabId, message) {
  await waitForMonitorTab(tabId);
  try {
    return await chrome.tabs.sendMessage(tabId,message,{frameId:0});
  } catch (_) {
    await chrome.scripting.executeScript({
      target:{tabId,frameIds:[0]},
      files:["teams-content.js"]
    });
    return chrome.tabs.sendMessage(tabId,message,{frameId:0});
  }
}

async function addCurrentGroup() {
  const tab=await activeTeamsTab();
  await chrome.scripting.executeScript({target:{tabId:tab.id,frameIds:[0]},files:["teams-content.js"]});
  const info=await chrome.tabs.sendMessage(tab.id,{type:"TEAMS_GET_GROUP_INFO_V17"},{frameId:0});
  if(!info?.ok || !info.groupName || !info.groupKey) {
    throw new Error(info?.error || "Could not identify the currently open Teams group.");
  }

  const groups=await getStoredGroups();
  let group=groups.find(item=>item.key===info.groupKey);
  if(group) {
    const tabIds=Array.isArray(group.tabIds)?group.tabIds:[];
    if(!tabIds.includes(tab.id)) tabIds.push(tab.id);
    group.tabIds=tabIds;
  } else {
    group={
      id:info.groupKey,key:info.groupKey,name:info.groupName,
      tabIds:[tab.id],lastAddedTabId:tab.id,priority:groups.length,readHistory:false
    };
    groups.push(group);
  }
  group.name=info.groupName;
  group.lastAddedTabId=tab.id;
  delete group.chatId;
  delete group.monitorUrl;
  delete group.monitorTabId;
  groups.forEach((item,index)=>item.priority=index);
  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  await setStatus("Added "+group.name+". Teams notifications from this group will be monitored.");
  return {ok:true,group,baselineCount:info.baselineCount||0};
}

async function removeGroup(groupKey) {
  const existing=await getStoredGroups();
  const removed=existing.find(group=>group.key===groupKey);
  const groups=existing.filter(group=>group.key!==groupKey);
  groups.forEach((group,index)=>group.priority=index);
  await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  if(Number.isInteger(removed?.monitorTabId)) {
    try { await chrome.tabs.remove(removed.monitorTabId); } catch (_) {}
  }
  await hydrateSpeechQueue();
  speechQueue=speechQueue.filter(item=>item.groupKey!==groupKey);
  await persistSpeechQueue();
  await setStatus("Group removed. Its background monitor and queued messages were cleared.");
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
  const tabId=group.lastAddedTabId;
  if(!Number.isInteger(tabId)) throw new Error("Open this Teams group and add it again before reading its history.");
  const result=await sendToMonitorTab(tabId,{type:"TEAMS_READ_LAST_FIVE_V16",groupKey});
  if(!result?.ok) throw new Error(result?.error||"Could not read the last five messages from the currently open group.");
  const messages=(result.messages||[]).slice(-5);
  if(!messages.length) {
    await setStatus("No readable incoming messages were found in "+group.name+".");
    return publishLiveState();
  }
  const items=[{
    groupKey:group.key,groupName:group.name,
    messages:messages.map(message=>({sender:message.sender,text:message.text,messageId:message.key})),
    historyBatch:true,priority:groups.indexOf(group),announceGroup:true,sequence:Date.now()
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
  speechQueue=(Array.isArray(saved.teamsSpeechQueue)?saved.teamsSpeechQueue:[])
    .map(normalizeSpeechBatch).filter(batch=>batch.messages.length>0);
  speechCurrent=saved.teamsSpeechCurrent?normalizeSpeechBatch(saved.teamsSpeechCurrent):null;
  speechQueueLoaded=true;

  if(speechCurrent?.phase==="gap") {
    const sequence=speechCurrent.sequence;
    groupDelayTimer=setTimeout(()=>{
      if(speechCurrent?.sequence===sequence&&speechCurrent.phase==="gap") {
        speechCurrent.phase="message";
        persistSpeechQueue().then(speakCurrentPhase);
      }
    },5000);
  }
}
function normalizeSpeechBatch(item){
  const rawMessages=Array.isArray(item?.messages)
    ? item.messages
    : (item?.sender||item?.audioRef||item?.audioUrl||item?.text)
      ? [{sender:item.sender||"",text:item.text||"",messageId:item.messageId||"",replyTo:item.replyTo||"",isReply:!!item.isReply,audioRef:item.audioRef||"",audioUrl:item.audioUrl||"",sourceTabId:item.sourceTabId}]
      : [];
  return {
    groupKey:String(item?.groupKey||""),
    groupName:String(item?.groupName||"Teams group"),
    messages:rawMessages.map(message=>({
      sender:String(message.sender||""),
      text:String(message.text||""),
      messageId:String(message.messageId||message.key||""),
      replyTo:String(message.replyTo||""),
      isReply:!!message.isReply,
      audioRef:String(message.audioRef||""),
      audioUrl:String(message.audioUrl||""),
      sourceTabId:Number.isInteger(message.sourceTabId)?message.sourceTabId:(Number.isInteger(item?.sourceTabId)?item.sourceTabId:null)
    })),
    priority:Number.isFinite(item?.priority)?item.priority:999,
    sequence:Number.isFinite(item?.sequence)?item.sequence:Date.now(),
    messageIndex:Number.isInteger(item?.messageIndex)?Math.max(0,item.messageIndex):0,
    phase:["gap","message","audio"].includes(item?.phase)?item.phase:"group",
    playbackId:String(item?.playbackId||"")
  };
}
function speechMessageKey(message){
  return message.messageId || [message.sender,message.text,message.audioRef,message.audioUrl].join("|");
}
async function persistSpeechQueue() {
  await chrome.storage.local.set({
    teamsSpeechQueue:speechQueue,
    teamsSpeechCurrent:speechCurrent
  });
}
function appendUniqueMessages(batch,messages){
  const known=new Set(batch.messages.map(speechMessageKey));
  for(const message of messages){
    const key=speechMessageKey(message);
    if(known.has(key)) continue;
    batch.messages.push(message);
    known.add(key);
  }
}
async function enqueueSpeech(items) {
  await hydrateSpeechQueue();
  for(const raw of items||[]){
    const batch=normalizeSpeechBatch(raw);
    if(!batch.messages.length) continue;
    let target=null;
    if(speechCurrent?.groupKey===batch.groupKey) target=speechCurrent;
    if(!target) target=speechQueue.find(item=>item.groupKey===batch.groupKey);
    if(target) {
      appendUniqueMessages(target,batch.messages);
    } else {
      batch.phase="group";
      batch.messageIndex=0;
      speechQueue.push(batch);
    }
  }
  await persistSpeechQueue();
  scheduleSpeechQueue();
}
function scheduleSpeechQueue() {
  if(speechCurrent||speechDispatchTimer||!speechQueue.length) return;
  speechDispatchTimer=setTimeout(()=>{
    speechDispatchTimer=null;
    processSpeechQueue();
  },850);
}
async function currentGroupPriority(groupKey) {
  const groups=await getStoredGroups();
  const index=groups.findIndex(group=>group.key===groupKey);
  return index<0?999:index;
}
async function processSpeechQueue() {
  if(speechProcessing) return;
  await hydrateSpeechQueue();
  if(speechCurrent||!speechQueue.length) return;
  speechProcessing=true;
  try{
    for(const item of speechQueue) item.priority=await currentGroupPriority(item.groupKey);
    speechQueue.sort((a,b)=>(a.priority-b.priority)||(a.sequence-b.sequence));
    speechCurrent=speechQueue.shift();
    speechCurrent.messageIndex=0;
    speechCurrent.phase="group";
    speechCurrent.playbackId="";
    await persistSpeechQueue();
    speakCurrentPhase();
  }finally{
    speechProcessing=false;
  }
}
function formattedMessage(message){
  const sender=message.sender||"Unknown sender";
  const body=String(message.text||"").trim();
  if(message.replyTo) return message.replyTo+" ke message ka reply "+sender+" ne diya"+(body?": "+body:"")+".";
  if(message.isReply) return sender+" ne ek purane message ka reply diya"+(body?": "+body:"")+".";
  return sender+" ne kaha"+(body?": "+body:"")+".";
}
function scheduleNextMessage(delay=350){
  if(!speechCurrent) return;
  const sequence=speechCurrent.sequence;
  if(speechAdvanceTimer) clearTimeout(speechAdvanceTimer);
  speechAdvanceTimer=setTimeout(()=>{
    speechAdvanceTimer=null;
    if(speechCurrent?.sequence!==sequence) return;
    speechCurrent.messageIndex++;
    if(speechCurrent.messageIndex<speechCurrent.messages.length){
      speechCurrent.phase="message";
      persistSpeechQueue().then(speakCurrentPhase);
    }else{
      finishCurrentSpeech();
    }
  },delay);
}
function startAudioPlayback(current,message){
  const tabId=message.sourceTabId;
  if(!Number.isInteger(tabId)){
    setStatus("Audio message detected, but its Teams player is not available in this tab.");
    scheduleNextMessage(350);
    return;
  }
  const playbackId=String(current.sequence)+":"+String(current.messageIndex)+":"+String(Date.now());
  current.phase="audio";
  current.playbackId=playbackId;
  persistSpeechQueue();
  chrome.tabs.sendMessage(tabId,{
    type:"TEAMS_PLAY_AUDIO",
    audioRef:message.audioRef||"",
    audioUrl:message.audioUrl||"",
    playbackId
  },response=>{
    if(chrome.runtime.lastError||!response?.ok){
      const error=chrome.runtime.lastError?.message||response?.error||"Audio player is not accessible.";
      setStatus("Could not play the Teams audio attachment: "+error);
      if(speechCurrent?.playbackId===playbackId) scheduleNextMessage(350);
      return;
    }
    setStatus("Playing audio attachment from "+current.groupName+". Speech is paused until audio ends.");
  });
}
function speakCurrentPhase() {
  if(!speechCurrent) return;
  const current=speechCurrent;
  if(current.phase==="gap"||current.phase==="audio") return;
  if(current.phase==="group"){
    speakText(current.groupName,{
      sequence:current.sequence,
      phase:"group"
    });
    return;
  }
  const message=current.messages[current.messageIndex];
  if(!message){finishCurrentSpeech();return;}
  if(message.audioRef||message.audioUrl){
    startAudioPlayback(current,message);
    return;
  }
  const text=formattedMessage(message);
  if(!text.replace(/[.:\s]/g,"")){scheduleNextMessage(100);return;}
  speakText(text,{sequence:current.sequence,phase:"message",sender:message.sender});
}
function speakText(text,context){
  const current=speechCurrent;
  if(!current||current.sequence!==context.sequence) return;
  const speakWithVoice=availableVoices=>{
    if(!speechCurrent||speechCurrent.sequence!==context.sequence) return;
    const voices=Array.isArray(availableVoices)?availableVoices:[];
    const voicesWithEnd=voices.filter(item=>Array.isArray(item.eventTypes)&&item.eventTypes.includes("end"));
    const voice=
      voicesWithEnd.find(item=>/^hi[-_]IN$/i.test(item.lang||""))||
      voicesWithEnd.find(item=>/^hi(?:[-_]|$)/i.test(item.lang||""))||
      voicesWithEnd.find(item=>/^en[-_]IN$/i.test(item.lang||""))||null;
    const options={
      enqueue:false,
      rate:context.phase==="group"?0.85:0.78,
      pitch:1,
      requiredEventTypes:["end"],
      desiredEventTypes:["start","end","error","interrupted","cancelled"],
      onEvent:event=>{
        if(!speechCurrent||speechCurrent.sequence!==context.sequence) return;
        if(event.type==="start"){
          setStatus(context.phase==="group"
            ?"Announcing group "+current.groupName
            :"Reading "+(context.sender||"sender")+" in "+current.groupName);
        }else if(event.type==="end"){
          if(context.phase==="group"){
            current.phase="gap";
            persistSpeechQueue().then(()=>{
              if(groupDelayTimer) clearTimeout(groupDelayTimer);
              groupDelayTimer=setTimeout(()=>{
                groupDelayTimer=null;
                if(speechCurrent?.sequence===context.sequence&&speechCurrent.phase==="gap"){
                  speechCurrent.phase="message";
                  persistSpeechQueue().then(speakCurrentPhase);
                }
              },5000);
            });
            setStatus("Group announced once. Waiting 5 seconds before its message batch.");
          }else if(current.phase==="message"){
            scheduleNextMessage(450);
          }
        }else if(["error","interrupted","cancelled"].includes(event.type)){
          if(event.type==="error") setStatus("Speech error: "+(event.errorMessage||"Chrome TTS could not read the message."));
          if(context.phase==="group") finishCurrentSpeech();
          else scheduleNextMessage(350);
        }
      }
    };
    if(voice?.voiceName){options.voiceName=voice.voiceName;options.lang=voice.lang;}
    else setStatus("No Hindi voice is installed. Trying the default Chrome speech voice.");
    chrome.tts.speak(text,options,()=>{
      if(chrome.runtime.lastError){
        setStatus("Chrome TTS error: "+chrome.runtime.lastError.message);
        if(context.phase==="group") finishCurrentSpeech();
        else scheduleNextMessage(350);
      }
    });
  };
  try{chrome.tts.getVoices(voices=>speakWithVoice(voices));}
  catch(error){
    setStatus("Could not load speech voices: "+error.message);
    speakWithVoice([]);
  }
}
async function finishCurrentSpeech() {
  speechCurrent=null;
  if(groupDelayTimer) clearTimeout(groupDelayTimer);
  if(speechAdvanceTimer) clearTimeout(speechAdvanceTimer);
  groupDelayTimer=null;
  speechAdvanceTimer=null;
  await persistSpeechQueue();
  scheduleSpeechQueue();
}
async function stopTeamsAudio(){
  for(const tab of await getTeamsTabs()){
    try{await chrome.tabs.sendMessage(tab.id,{type:"TEAMS_STOP_AUDIO"});}catch(_){}
  }
}
async function handleAudioPlaybackEvent(message){
  if(!speechCurrent||speechCurrent.phase!=="audio"||speechCurrent.playbackId!==message.playbackId) return;
  if(message.type==="TEAMS_AUDIO_FAILED") await setStatus(message.error||"Teams audio playback failed.");
  scheduleNextMessage(350);
}
async function onTeamsMessage(message,sender) {
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(!enabled[TEAMS_LIVE_KEY]) return;
  const groupKey=String(message.groupKey||"");
  const groups=await getStoredGroups();
  const group=groups.find(item=>item.key===groupKey);
  if(!group) return;

  const tabId=sender?.tab?.id;
  const sourceUrl=sender?.url||sender?.tab?.url||"";
  if(!Number.isInteger(tabId)||!teamsUrl(sourceUrl)) return;

  const senderName=String(message.sender||"").trim();
  const text=String(message.text||"").trim();
  const audioRef=String(message.audioRef||"");
  const audioUrl=String(message.audioUrl||"");
  if(!senderName||(!text&&!audioRef&&!audioUrl)) return;

  const messageId=String(message.messageId||"");
  const stableKey=groupKey+"|"+(messageId||[senderName,text,audioRef,audioUrl,message.timestamp||""].join("|"));
  const now=Date.now();
  const recentTextKey=groupKey+"|"+senderName+"|"+text.toLocaleLowerCase()+"|"+(audioRef||audioUrl);
  const recentAt=recentMessageIds.get(recentTextKey)||0;
  if(recentAt&&now-recentAt<5000&&messageId.startsWith("teams-toast:")) return;
  if(recentMessageIds.has(stableKey)) return;
  recentMessageIds.set(stableKey,now);
  recentMessageIds.set(recentTextKey,now);
  if(recentMessageIds.size>5000){
    for(const [key,time] of recentMessageIds) if(now-time>120000) recentMessageIds.delete(key);
    if(recentMessageIds.size>5000) recentMessageIds.clear();
  }

  const entry={
    groupName:group.name,groupKey,sender:senderName,text,
    replyTo:String(message.replyTo||""),
    isReply:!!message.isReply,
    audioRef,audioUrl,
    sourceTabId:tabId,
    messageId:messageId||stableKey,
    receivedAt:now
  };
  const saved=await chrome.storage.local.get([TEAMS_READ_COUNT,TEAMS_RECENT_EVENTS]);
  const recent=[entry,...(saved[TEAMS_RECENT_EVENTS]||[])].slice(0,8);
  await chrome.storage.local.set({
    [TEAMS_READ_COUNT]:(saved[TEAMS_READ_COUNT]||0)+1,
    [TEAMS_RECENT_EVENTS]:recent,
    teamsLastStatus:audioRef||audioUrl
      ?"Audio message detected in "+group.name+"."
      :"New message in "+group.name+" from "+senderName+"."
  });
  await enqueueSpeech([{
    groupKey,groupName:group.name,
    messages:[{
      sender:senderName,text,messageId:entry.messageId,
      replyTo:entry.replyTo,isReply:entry.isReply,
      audioRef,audioUrl,sourceTabId:tabId
    }],
    priority:groups.indexOf(group),
    sequence:now
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
  if(message?.type==="TEAMS_AUDIO_ENDED"||message?.type==="TEAMS_AUDIO_FAILED") handleAudioPlaybackEvent(message).catch(()=>{});
  if(message?.type==="TEAMS_LIVE_STATUS") {
    const text=String(message.text||"");
    chrome.storage.local.set({teamsLastStatus:text});
  }
});


chrome.tabs.onUpdated.addListener(async(tabId,changeInfo,tab)=>{
  if(changeInfo.status!=="complete" || !teamsUrl(tab?.url)) return;
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY]) {
    try {
      await chrome.scripting.executeScript({target:{tabId,frameIds:[0]},files:["teams-content.js"]});
      await setStatus("Teams notification observer attached to the open Teams tab.");
    } catch(error) {
      await setStatus("Could not attach Teams notification observer: "+error.message);
    }
  }
});
chrome.tabs.onRemoved.addListener(async tabId=>{
  const groups=await getStoredGroups();
  let changed=false;
  for(const group of groups) {
    const before=(group.tabIds||[]).length;
    group.tabIds=(group.tabIds||[]).filter(id=>id!==tabId);
    if(before!==group.tabIds.length) changed=true;
    if(group.lastAddedTabId===tabId){group.lastAddedTabId=null;changed=true;}
    if(group.monitorTabId===tabId){group.monitorTabId=null;changed=true;}
  }
  if(changed) await chrome.storage.local.set({[TEAMS_GROUPS_KEY]:groups});
  publishLiveState();
});
chrome.runtime.onStartup.addListener(async()=>{
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY]) {
    const result=await injectTeamsTabs();
    await setStatus("Teams notification observers restored: "+result.injected+"/"+result.total+" tab(s).");
  }
});
chrome.runtime.onInstalled.addListener(async()=>{
  const enabled=await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if(enabled[TEAMS_LIVE_KEY]) {
    const result=await injectTeamsTabs();
    await setStatus("Teams notification observers attached: "+result.injected+"/"+result.total+" tab(s).");
  }
});

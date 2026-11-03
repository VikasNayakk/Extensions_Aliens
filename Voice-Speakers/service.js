import {localAnalyze,redactSensitiveText,SYSTEM_PROMPT} from "./analyzer.js";

chrome.runtime.onMessage.addListener((m,s,send)=>{if(m?.type==="ANALYZE_MESSAGES"){analyze(m.text).then(send).catch(e=>send({output:"Analysis error: "+e.message,remote:false}));return true;}if(m?.type==="TEST_OPENAI"){testApi().then(send).catch(e=>send({ok:false,error:e.message}));return true;}});
async function requestOpenAI(key,model,messages){let r;try{r=await fetch("https://api.openai.com/v1/chat/completions",{method:"POST",headers:{"Authorization":"Bearer "+key,"Content-Type":"application/json"},body:JSON.stringify({model,temperature:0,messages})});}catch(e){throw Error("Network/CORS error: "+e.message);}const data=await r.json().catch(()=>({}));if(!r.ok){const detail=data?.error?.message||data?.error?.code||("HTTP "+r.status);throw Error("OpenAI API "+r.status+": "+detail);}const output=data.choices?.[0]?.message?.content?.trim();if(!output)throw Error("OpenAI returned an empty response.");return output;}
async function testApi(){const s=await chrome.storage.local.get(["apiKey","model"]);if(!s.apiKey)throw Error("Save an API key first.");await requestOpenAI(s.apiKey,s.model||"gpt-4.1-mini",[{role:"user",content:"Reply with exactly: API connection successful"}]);return {ok:true,model:s.model||"gpt-4.1-mini"};}
async function analyze(input){const s=await chrome.storage.local.get(["apiKey","model","redact"]);const safe=s.redact===false?String(input):redactSensitiveText(input);if(!s.apiKey)return {output:localAnalyze(safe),remote:false};const output=await requestOpenAI(s.apiKey,s.model||"gpt-4.1-mini",[{role:"system",content:SYSTEM_PROMPT},{role:"user",content:safe}]);return {output,remote:true};}

const TEAMS_LIVE_KEY = "teamsLiveEnabled";
const recentMessageIds = new Set();

async function getTeamsTabs() {
  const tabs = await chrome.tabs.query({});
  return tabs.filter(tab => /^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(tab.url || ""));
}

async function getLiveState() {
  const saved = await chrome.storage.local.get([
    TEAMS_LIVE_KEY,
    "teamsReadCount",
    "teamsRecentEvents",
    "teamsLastStatus"
  ]);
  const tabs = await getTeamsTabs();
  return {
    teamsLiveEnabled: !!saved[TEAMS_LIVE_KEY],
    teamsTabCount: tabs.length,
    teamsReadCount: saved.teamsReadCount || 0,
    teamsRecentEvents: saved.teamsRecentEvents || [],
    teamsLastStatus: saved.teamsLastStatus || ""
  };
}

async function publishLiveState() {
  const state = await getLiveState();
  chrome.runtime.sendMessage({type:"TEAMS_LIVE_STATE_UPDATED",state},()=>void chrome.runtime.lastError);
  return state;
}

async function injectTeamsTabs() {
  const tabs = await getTeamsTabs();
  let injected = 0;
  let failed = 0;
  for (const tab of tabs) {
    try {
      await chrome.scripting.executeScript({
        target:{tabId:tab.id,allFrames:true},
        files:["teams-content.js"]
      });
      injected++;
    } catch (_) {
      failed++;
    }
  }
  return {total:tabs.length,injected,failed};
}

async function setTeamsLiveEnabled(enabled) {
  await chrome.storage.local.set({[TEAMS_LIVE_KEY]:!!enabled});
  if (!enabled) {
    chrome.tts.stop();
    await chrome.storage.local.set({teamsLastStatus:"Live reading is OFF."});
  } else {
    const result = await injectTeamsTabs();
    const status = result.total
      ? "Live reading is ON. " + result.total + " Teams tab(s) found; existing messages are being skipped."
      : "Live reading is ON. No Teams tab is open yet; monitoring will start automatically when Teams opens.";
    await chrome.storage.local.set({teamsLastStatus:status});
  }
  return publishLiveState();
}

async function onTeamsMessage(message, sender) {
  const settings = await chrome.storage.local.get([TEAMS_LIVE_KEY,"teamsReadCount","teamsRecentEvents"]);
  if (!settings[TEAMS_LIVE_KEY]) return;
  const senderName = String(message.sender || "").trim();
  const text = String(message.text || "").trim();
  const messageId = String(message.messageId || "");
  if (!senderName || !text) return;
  const uniqueId = messageId || ((sender?.tab?.id || "tab") + "|" + (sender?.frameId || 0) + "|" + senderName + "|" + text);
  if (messageId && recentMessageIds.has(uniqueId)) return;
  if (messageId) {
    recentMessageIds.add(uniqueId);
    if (recentMessageIds.size > 2000) recentMessageIds.clear();
  }

  const entry = {sender:senderName,text,receivedAt:Date.now()};
  const recent = [entry,...(settings.teamsRecentEvents || [])].slice(0,8);
  const count = (settings.teamsReadCount || 0) + 1;
  await chrome.storage.local.set({
    teamsReadCount:count,
    teamsRecentEvents:recent,
    teamsLastStatus:"Reading new message from " + senderName + (sender?.tab?.id ? " · Teams tab " + sender.tab.id : "")
  });

  chrome.tts.speak(senderName + " ne kaha: " + text, {
    enqueue:true,
    rate:0.85,
    lang:"hi-IN"
  }, () => {
    if (chrome.runtime.lastError) {
      chrome.storage.local.set({teamsLastStatus:"Speech error: " + chrome.runtime.lastError.message});
    }
  });
  await publishLiveState();
}

chrome.runtime.onMessage.addListener((message,sender,sendResponse) => {
  if (message?.type === "GET_TEAMS_LIVE_STATUS") {
    getLiveState().then(sendResponse).catch(error=>sendResponse({teamsLastStatus:error.message}));
    return true;
  }
  if (message?.type === "SET_TEAMS_LIVE") {
    setTeamsLiveEnabled(!!message.enabled).then(sendResponse).catch(error=>sendResponse({teamsLastStatus:error.message}));
    return true;
  }
  if (message?.type === "TEAMS_LIVE_MESSAGE") {
    onTeamsMessage(message,sender).catch(()=>{});
  }
  if (message?.type === "TEAMS_LIVE_STATUS") {
    const tabId = sender?.tab?.id;
    const text = String(message.text || "");
    if (tabId) {
      chrome.storage.local.get(["teamsLiveFrames"]).then(saved => {
        const frames = {...(saved.teamsLiveFrames || {}),[tabId]:{text,at:Date.now()}};
        chrome.storage.local.set({teamsLiveFrames:frames,teamsLastStatus:text});
      });
    }
  }
});

chrome.tabs.onUpdated.addListener(async (tabId,changeInfo,tab) => {
  if (changeInfo.status !== "complete" || !/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(tab?.url || "")) return;
  const saved = await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if (!saved[TEAMS_LIVE_KEY]) return;
  try {
    await chrome.scripting.executeScript({target:{tabId,allFrames:true},files:["teams-content.js"]});
  } catch (_) {}
  await publishLiveState();
});

chrome.tabs.onCreated.addListener(async tab => {
  if (!/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(tab.url || "")) return;
  const saved = await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if (saved[TEAMS_LIVE_KEY] && tab.id) {
    try {
      await chrome.scripting.executeScript({target:{tabId:tab.id,allFrames:true},files:["teams-content.js"]});
    } catch (_) {}
  }
});

chrome.tabs.onRemoved.addListener(() => publishLiveState());

chrome.runtime.onStartup.addListener(async () => {
  const saved = await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if (saved[TEAMS_LIVE_KEY]) await injectTeamsTabs();
});
chrome.runtime.onInstalled.addListener(async () => {
  const saved = await chrome.storage.local.get(TEAMS_LIVE_KEY);
  if (saved[TEAMS_LIVE_KEY]) await injectTeamsTabs();
});

(() => {
  const STATE_KEY="__aliensVoiceSpeakersTeamsContentV15";
  if(window[STATE_KEY]) return;

  const state=window[STATE_KEY]={
    running:false,enabled:false,root:null,rootSelector:"",chatKey:"",
    seen:new Set(),observer:null,waitingObserver:null,poller:null,timer:null,suppressUntil:0
  };

  const send=(type,payload={})=>{
    try{chrome.runtime.sendMessage({type,...payload},()=>void chrome.runtime.lastError);}catch(_){}
  };
  const clean=value=>String(value||"").replace(/\u200b/g,"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
  const normalize=value=>clean(value).toLocaleLowerCase();
  const visible=node=>!!node&&node.getClientRects().length>0;

  const rootSelectors=[
    '[data-tid="thread-body-scrollable-content"]',
    '[data-tid="message-pane-body"]',
    '[data-tid="chat-pane-message-list"]',
    '[data-tid="scrollable-thread-body"]',
    '[aria-label*="Message List" i]',
    'div[role="log"]',
    'div[role="list"][aria-label*="message" i]',
    '[data-tid="thread-view"]',
    '[data-tid="message-pane"]'
  ];
  const bodySelectors=[
    '[id^="message-body-"]',
    '[data-tid="message-content"]',
    '[data-tid="message-body"]',
    '[data-tid="bubble-text"]',
    '[data-tid="chat-pane-message-text"]',
    '[data-tid="message-bubble"]',
    '[data-tid^="message-content-"]',
    '[data-tid^="messageBody"]',
    ".ts-message-content",
    ".message-body",
    '[aria-label^="Message by "]'
  ];
  const itemSelectors=[
    '[data-tid="message-group-container"]',
    '[data-tid="message-list-item"]',
    '[data-tid="message-list-item-container"]',
    'div[role="listitem"]',
    '[aria-label^="Message by "]'
  ];
  const authorSelectors=[
    '[data-tid="message-author-name"]',
    '[data-tid="author-name"]',
    '[id^="message-author-"]',
    '[data-tid="message-group-author"]',
    '[data-tid="message-group-header"]'
  ];
  const timeSelectors=[
    '[id^="timestamp-"]',
    '[data-tid="message-time"]',
    '[data-tid="message-timestamp"]',
    "time[datetime]",
    "time"
  ];

  function groupName(){
    const title=clean(document.title);
    const match=title.match(/^(?:\(\d+\)\s*)?Chat\s*\|\s*(.*?)\s*\|\s*Microsoft Teams$/i);
    if(match&&clean(match[1])) return clean(match[1]);
    const selectors=[
      '[data-tid="entity-header"]',
      '[data-tid="thread-header"]',
      '[data-tid="chat-header"]',
      '[data-tid="chat-title"]',
      '[data-tid="chat-name"]',
      '[data-tid="conversation-title"]',
      'header [role="heading"]',
      'h1'
    ];
    for(const selector of selectors){
      const node=document.querySelector(selector);
      if(node){
        const value=clean(node.innerText||node.textContent);
        if(value && !/^(chat|message list)$/i.test(value)) return value.split("\n")[0];
      }
    }
    return "";
  }
  function groupKey(){return normalize(groupName());}

  function countBodies(root){
    let count=0;
    for(const selector of bodySelectors) count+=root.querySelectorAll(selector).length;
    if(root.matches?.('[aria-label^="Message by "]')) count++;
    return count;
  }
  function findRoot(){
    const choices=[];
    for(const selector of rootSelectors){
      for(const node of document.querySelectorAll(selector)){
        if(!visible(node)) continue;
        const score=countBodies(node);
        if(score>0) choices.push({node,score,selector});
      }
    }
    choices.sort((a,b)=>b.score-a.score);
    return choices[0]||null;
  }
  function bodiesIn(node){
    for(const selector of bodySelectors){
      const found=[];
      if(node.matches?.(selector)&&visible(node)) found.push(node);
      found.push(...Array.from(node.querySelectorAll(selector)).filter(visible));
      if(found.length) return found.filter((body,index)=>!found.some((other,j)=>j!==index&&other.contains(body)));
    }
    return [];
  }
  function findItems(root){
    for(const selector of itemSelectors){
      const found=Array.from(root.querySelectorAll(selector)).filter(item=>visible(item)&&bodiesIn(item).length);
      if(found.length) return found.filter(item=>!found.some(other=>other!==item&&other.contains(item)));
    }
    return bodiesIn(root).map(body=>body.closest(itemSelectors.join(","))||body.parentElement)
      .filter((item,index,all)=>item&&all.indexOf(item)===index);
  }
  function meta(start,selectors){
    for(let node=start,depth=0;node&&depth<7;node=node.parentElement,depth++){
      for(const selector of selectors){
        const found=node.matches?.(selector)?node:node.querySelector(selector);
        if(found){
          const value=clean(found.getAttribute("aria-label")||found.getAttribute("title")||found.innerText||found.textContent);
          if(value) return value;
        }
      }
      const label=node.getAttribute?.("aria-label")||"";
      const match=label.match(/^Message by\s+([^,]+)/i);
      if(match) return clean(match[1]);
    }
    return "";
  }
  function ownName(){
    for(const node of document.querySelectorAll("[data-tid], [role='treeitem'], [role='listitem']")){
      if(!visible(node)) continue;
      const lines=String(node.innerText||node.textContent||"").split(/\r?\n/).map(clean).filter(Boolean);
      const own=lines.find(line=>/^.{2,80}?\s*\(You\)$/i.test(line));
      if(own) return clean(own.replace(/\s*\(You\)$/i,""));
    }
    return "";
  }
  function redact(text){
    return clean(text)
      .replace(/\b(password|passcode|passwd|pwd|otp|one[- ]time code|api[ _-]?key|access[ _-]?token|secret|2fa)\b\s*(?:is\s*)?[:=]?\s*([^\s,;]+)/gi,"$1 [hidden]")
      .replace(/\bBearer\s+[A-Za-z0-9._~+\/-]+=*/gi,"Bearer [hidden]")
      .replace(/\b(?:sk-[A-Za-z0-9_-]{16,}|gh[pousr]_[A-Za-z0-9]{20,})\b/g,"[access token hidden]")
      .replace(/\b(?=[A-Za-z0-9!@#$%^&*_-]{12,}\b)(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[!@#$%^&*_-])[A-Za-z0-9!@#$%^&*_-]+\b/g,"[possible password hidden]");
  }
  function isOwn(item,body,sender,me){
    if(/^(you|me)$/i.test(sender)) return true;
    if(me&&sender.toLowerCase()===me.toLowerCase()) return true;
    for(let node=body,depth=0;node&&depth<8;node=node.parentElement,depth++){
      const value=[
        typeof node.className==="string"?node.className:"",
        node.getAttribute("data-tid"),
        node.getAttribute("data-message-direction"),
        node.getAttribute("data-is-own-message"),
        node.getAttribute("aria-label")
      ].filter(Boolean).join(" ").toLowerCase();
      if(/outgoing|message-out|own-message|self-message|sent-by-me|from-me|you sent|you said/.test(value)) return true;
      if(node===item) break;
    }
    return false;
  }
  function extract(root){
    const items=findItems(root),messages=[],me=ownName();
    items.forEach((item,index)=>{
      const groupAuthor=meta(item,authorSelectors);
      bodiesIn(item).forEach(body=>{
        const label=body.getAttribute("aria-label")||item.getAttribute("aria-label")||"";
        const labelMatch=label.match(/^Message by\s+([^,]+)/i);
        const sender=clean(meta(body,authorSelectors)||groupAuthor||(labelMatch&&labelMatch[1])||"");
        if(!sender) return;
        const source=body.querySelector('[id^="content-"]')||body;
        const clone=source.cloneNode(true);
        clone.querySelectorAll("button,[role='button'],svg,script,style,nav,[aria-hidden='true'],[data-tid*='reaction'],[data-testid*='reaction']").forEach(node=>node.remove());
        const raw=clean(clone.innerText||clone.textContent);
        if(!raw||/^Message by\s+/i.test(raw)) return;
        const stamp=meta(body,timeSelectors)||meta(item,timeSelectors);
        const id=body.id||body.getAttribute("data-message-id")||item.getAttribute("data-message-id")||"";
        const key=id||[sender,stamp,raw].join("|");
        messages.push({
          key,sender,text:redact(raw),timestamp:stamp,own:isOwn(item,body,sender,me),index
        });
      });
    });
    return {messages,itemCount:items.length};
  }
  function chatInfo(){
    const name=groupName(),key=normalize(name);
    const root=findRoot();
    const data=root?extract(root.node):{messages:[],itemCount:0};
    return {ok:!!name&&!!root,groupName:name,groupKey:key,baselineCount:data.messages.length,itemCount:data.itemCount,error:!name?"Teams group name could not be detected.":!root?"Open the group chat so its messages are visible.":""};
  }
  function stop(){
    state.running=false;
    state.enabled=false;
    state.observer?.disconnect();
    state.waitingObserver?.disconnect();
    if(state.timer) clearTimeout(state.timer);
    if(state.poller) clearInterval(state.poller);
    state.observer=state.waitingObserver=state.poller=null;
    state.root=null;
  }
  function start(){
    if(!/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(location.href)||state.running) return;
    state.running=state.enabled=true;
    state.root=null;
    state.seen.clear();
    function schedule(){
      if(!state.running||state.timer) return;
      state.timer=setTimeout(()=>{state.timer=null;scan();},180);
    }
    function attach(info){
      state.root=info.node;
      state.rootSelector=info.selector;
      state.chatKey=groupKey();
      const baseline=extract(info.node);
      state.seen=new Set(baseline.messages.map(message=>message.key));
      state.observer?.disconnect();
      state.observer=new MutationObserver(schedule);
      state.observer.observe(info.node,{childList:true,subtree:true,characterData:true});
      send("TEAMS_LIVE_STATUS",{running:true,text:"Group tab ready: "+(groupName()||"Unknown group")+". Existing messages skipped: "+baseline.messages.length+"."});
    }
    function scan(){
      if(!state.running) return;
      const info=findRoot();
      if(!info) return;
      if(info.node!==state.root||info.selector!==state.rootSelector||groupKey()!==state.chatKey){
        attach(info);
        return;
      }
      const batch=extract(info.node);
      if(Date.now()<state.suppressUntil){
        state.seen=new Set(batch.messages.map(message=>message.key));
        return;
      }
      const unseen=batch.messages.filter(message=>!state.seen.has(message.key));
      // Only messages inserted at the end of the open conversation count as new.
      const newAtEnd=unseen.filter(message=>message.index>=Math.max(0,batch.messages.length-unseen.length-5));
      unseen.forEach(message=>state.seen.add(message.key));
      newAtEnd.forEach(message=>{
        if(message.own) return;
        send("TEAMS_LIVE_MESSAGE",{
          groupName:groupName(),
          groupKey:groupKey(),
          sender:message.sender,
          text:message.text,
          timestamp:message.timestamp,
          messageId:message.key
        });
      });
      if(state.seen.size>3000) state.seen=new Set(Array.from(state.seen).slice(-1800));
    }
    state.waitingObserver=new MutationObserver(schedule);
    state.waitingObserver.observe(document.documentElement,{childList:true,subtree:true});
    scan();
    state.poller=setInterval(scan,1200);
    send("TEAMS_LIVE_STATUS",{running:true,text:"Teams live reader ON. Waiting for selected group updates."});
  }

  chrome.runtime.onMessage.addListener((message,sender,sendResponse)=>{
    if(message?.type==="TEAMS_GET_CHAT_INFO"){
      const info=chatInfo();
      sendResponse(info);
      return true;
    }
    if(message?.type==="TEAMS_READ_LAST_FIVE"){
      const info=chatInfo();
      if(!info.ok||info.groupKey!==message.groupKey){
        sendResponse({ok:false,error:"The current Teams tab is no longer on this selected group. Open the group and add it again."});
        return true;
      }
      const root=findRoot();
      const batch=root?extract(root.node):{messages:[]};
      const messages=batch.messages.filter(message=>!message.own).slice(-5);
      sendResponse({ok:true,groupName:info.groupName,messages});
      return true;
    }
  });

  chrome.storage.onChanged.addListener((changes,area)=>{
    if(area==="local"&&changes.teamsLiveEnabled){
      if(changes.teamsLiveEnabled.newValue) start();
      else stop();
    }
  });
  chrome.storage.local.get("teamsLiveEnabled").then(saved=>{if(saved.teamsLiveEnabled) start();}).catch(()=>{});
})();
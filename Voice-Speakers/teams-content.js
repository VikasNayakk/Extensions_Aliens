(() => {
  const STATE_KEY="__aliensVoiceSpeakersTeamsContentV16";
  if(window[STATE_KEY]) return;

  const state=window[STATE_KEY]={
    running:false,enabled:false,root:null,rootSelector:"",chatKey:"",
    seen:new Set(),toastSeen:new Map(),toastPrimed:false,lastSeenKey:null,hasBaseline:false,observer:null,waitingObserver:null,
    toastObserver:null,toastPoller:null,poller:null,timer:null,suppressUntil:0,activeAudio:null
  };

  const send=(type,payload={})=>{
    try{chrome.runtime.sendMessage({type,...payload},()=>void chrome.runtime.lastError);}catch(_){}
  };
  const clean=value=>String(value||"").replace(/\u200b/g,"").replace(/\u00a0/g," ").replace(/\s+/g," ").trim();
  const normalize=value=>clean(value).normalize("NFKC").toLocaleLowerCase();
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
  function conversationChatId(){
    const selectors=[
      'a[href*="/l/message/"]',
      'a[href*="/l/chat/"]'
    ];
    for(const selector of selectors){
      for(const link of document.querySelectorAll(selector)){
        const raw=link.getAttribute("href")||link.href||"";
        if(!raw) continue;
        try{
          const url=new URL(raw,location.origin);
          let match=url.pathname.match(/\/l\/message\/([^/]+)/i);
          if(!match) match=url.pathname.match(/\/l\/chat\/([^/]+)/i);
          if(match){
            const id=decodeURIComponent(match[1]);
            if(/^19:/i.test(id)) return id;
          }
        }catch(_){}
      }
    }

    const attributeSelectors=["[data-conversation-id]","[data-chat-id]","[data-thread-id]"];
    for(const selector of attributeSelectors){
      for(const node of document.querySelectorAll(selector)){
        for(const attribute of ["data-conversation-id","data-chat-id","data-thread-id"]){
          const value=node.getAttribute(attribute)||"";
          if(/^19:/i.test(value.trim())) return value.trim();
        }
      }
    }
    return "";
  }
  function monitorUrl(chatId){
    if(!chatId) return "";
    return "https://teams.microsoft.com/l/chat/"+encodeURIComponent(chatId)+"/conversations";
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
  function isTimestampLine(line){
    const value=clean(line).toLowerCase();
    return /^(?:(?:today|yesterday|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s*,?\s*)?\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?$/.test(value) ||
      /^(?:today|yesterday)\s+\d{1,2}(?::\d{2})?\s*(?:am|pm)?$/.test(value) ||
      /^\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?$/.test(value) ||
      /^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2}(?:,?\s+\d{1,2}:\d{2}\s*(?:am|pm)?)?$/i.test(value);
  }
  function speechText(value){
    const lines=String(value||"").split(/\r?\n/).map(clean).filter(line=>line&&!isTimestampLine(line));
    return lines.join(" ")
      .replace(/[\p{Extended_Pictographic}\uFE0F\u200D]/gu," ")
      .replace(/[\u{1F1E6}-\u{1F1FF}]/gu," ")
      .replace(/[#*0-9]\uFE0F?\u20E3/gu," ")
      .replace(/\s+/g," ").trim();
  }
  function replyMeta(item,body){
    const selectors=[
      '[data-tid*="quoted-message" i]',
      '[data-tid*="message-reply" i]',
      '[data-tid*="reply-preview" i]',
      '[data-testid*="quoted-message" i]',
      '[data-tid*="quoted" i]',
      '[aria-label*="replying to" i]',
      '[aria-label*="quoted message" i]'
    ];
    const nodes=[];
    for(const selector of selectors) nodes.push(...item.querySelectorAll(selector));
    const unique=nodes.filter((node,index)=>nodes.indexOf(node)===index);
    for(const node of unique){
      const label=clean(node.getAttribute("aria-label")||node.getAttribute("title")||"");
      const labelMatch=label.match(/(?:replying to|reply to|quoted message from|in reply to)\s*[:\-]?\s*(.+)$/i);
      const authorNode=node.querySelector('[data-tid*="author" i],[data-testid*="author" i],[class*="author" i],strong,b');
      const author=clean(authorNode?.innerText||authorNode?.textContent||labelMatch?.[1]||"");
      if(author){
        const quotedNode=node.querySelector("[data-message-id],[data-mid]")||node;
        const quotedId=quotedNode.getAttribute("data-message-id")||quotedNode.getAttribute("data-mid")||"";
        return {
          replyTo:author.replace(/^(?:replying to|reply to|quoted message from)\s*[:\-]?\s*/i,"").split(/\r?\n/)[0],
          replyToMessageId:quotedId,
          nodes:unique
        };
      }
    }
    const quotedNode=unique.find(node=>node.hasAttribute("data-message-id")||node.hasAttribute("data-mid"));
    return {
      replyTo:"",
      replyToMessageId:quotedNode?(quotedNode.getAttribute("data-message-id")||quotedNode.getAttribute("data-mid")||""):"",
      nodes:unique
    };
  }
  function audioAttachment(item,body,messageKey){
    const audio=item.querySelector("audio")||body.querySelector("audio");
    if(audio){
      let ref=audio.getAttribute("data-vs-audio-ref");
      if(!ref){
        ref="vs-audio-"+String(messageKey||Math.random().toString(36).slice(2)).replace(/[^a-z0-9_-]/gi,"").slice(0,90);
        audio.setAttribute("data-vs-audio-ref",ref);
      }
      return {audioRef:ref};
    }
    const audioPattern=/\.(?:mp3|m4a|aac|wav|ogg|opus|flac|webm)(?:[?#]|$)/i;
    const candidates=Array.from(item.querySelectorAll("a[href], [data-url], [data-href], [data-download-url]"));
    for(const node of candidates){
      const raw=node.getAttribute("href")||node.getAttribute("data-url")||node.getAttribute("data-href")||node.getAttribute("data-download-url")||"";
      const label=clean(node.getAttribute("aria-label")||node.getAttribute("title")||node.innerText||node.textContent);
      if(!raw||(!audioPattern.test(raw)&&!audioPattern.test(label))) continue;
      try{
        const url=new URL(raw,location.href);
        if(/^https?:$/.test(url.protocol)||url.protocol==="blob:") return {audioUrl:url.href};
      }catch(_){}
    }
    return null;
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
        const messageId=body.id||body.getAttribute("data-message-id")||item.getAttribute("data-message-id")||"";
        const key=messageId||[sender,clean(meta(body,timeSelectors)),speechText(source.innerText||source.textContent)].join("|");
        const media=audioAttachment(item,body,key);
        const reply=replyMeta(item,body);
        const clone=source.cloneNode(true);
        clone.querySelectorAll("button,[role='button'],svg,script,style,nav,[aria-hidden='true'],[data-tid*='reaction'],[data-testid*='reaction']").forEach(node=>node.remove());
        reply.nodes.forEach(node=>{
          if(clone.contains(node)) node.remove();
          else {
            const matching=clone.querySelector('[data-tid*="quoted-message" i],[data-tid*="message-reply" i],[data-tid*="reply-preview" i],[aria-label*="replying to" i],[aria-label*="quoted message" i]');
            matching?.remove();
          }
        });
        const raw=speechText(clone.innerText||clone.textContent);
        if(!raw&&!media) return;
        const stamp=meta(body,timeSelectors)||meta(item,timeSelectors);
        messages.push({
          key,sender,text:redact(raw),timestamp:stamp,
          own:isOwn(item,body,sender,me),index:messages.length,
          replyTo:reply.replyTo||"",
          replyToMessageId:reply.replyToMessageId||"",
          isReply:reply.nodes.length>0,
          ...(media||{})
        });
      });
    });
    return {messages,itemCount:items.length};
  }
  function chatInfo(){
    const name=groupName(),key=normalize(name),chatId=conversationChatId();
    const root=findRoot();
    const data=root?extract(root.node):{messages:[],itemCount:0};
    const url=monitorUrl(chatId);
    const error=!name
      ?"Teams group name could not be detected."
      :!root
        ?"Open the group chat so its messages are visible."
        :!chatId
          ?"Teams chat ID could not be read from message links. Open a message's More options > Copy link once, then retry."
          :"";
    return {
      ok:!!name&&!!root&&!!chatId,
      groupName:name,
      groupKey:key,
      chatId,
      monitorUrl:url,
      baselineCount:data.messages.length,
      itemCount:data.itemCount,
      error
    };
  }
  function stop(){
    state.running=false;
    state.enabled=false;
    state.observer?.disconnect();
    state.waitingObserver?.disconnect();
    state.toastObserver?.disconnect();
    if(state.activeAudio){try{state.activeAudio.pause();state.activeAudio.currentTime=0;}catch(_){}state.activeAudio=null;}
    if(state.timer) clearTimeout(state.timer);
    if(state.poller) clearInterval(state.poller);
    if(state.toastPoller) clearInterval(state.toastPoller);
    state.observer=state.waitingObserver=state.toastObserver=state.poller=state.toastPoller=null;
    state.root=null;
  }
  const toastBoilerplate=/^(microsoft teams|more options|send reaction|dismiss notification|reply|profile picture|close|notification)$/i;
  function toastCard(button){
    for(let node=button,depth=0;node&&node!==document.body&&depth<10;node=node.parentElement,depth++){
      const lines=String(node.innerText||node.textContent||"").split(/\r?\n/).map(clean).filter(Boolean);
      const senderIndex=lines.findIndex(line=>/^[^:]{1,90}:\s*\S/.test(line));
      if(senderIndex>0 && lines.length<=45){
        const group=lines.slice(0,senderIndex).reverse().find(line=>!toastBoilerplate.test(line)&&!/^\d{1,2}[:/]\d/.test(line));
        if(group) return node;
      }
    }
    return null;
  }
  function parseToast(card){
    const lines=String(card.innerText||card.textContent||"").split(/\r?\n/).map(clean).filter(Boolean);
    const senderIndex=lines.findIndex(line=>/^[^:]{1,90}:\s*\S/.test(line));
    if(senderIndex<=0) return null;
    const groupName=lines.slice(0,senderIndex).reverse().find(line=>!toastBoilerplate.test(line)&&!/^\d{1,2}[:/]\d/.test(line));
    if(!groupName) return null;
    const pair=lines[senderIndex].match(/^([^:]{1,90}):\s*(.*)$/);
    if(!pair) return null;
    let text=clean(pair[2]);
    if(!text && lines[senderIndex+1]) text=clean(lines[senderIndex+1]);
    if(!text) return null;
    const timeLabel=lines.find(line=>isTimestampLine(line))||"";
    const media=audioAttachment(card,card,[groupName,pair[1],timeLabel,text].join("|"));
    return {groupName,sender:clean(pair[1]),text:speechText(redact(text)),timeLabel,...(media||{})};
  }
  function scanToastNotifications(){
    if(!state.running||window.top!==window) return;
    document.querySelectorAll('[aria-label="Dismiss notification"]').forEach(button=>{
      const card=toastCard(button);
      if(!card) return;
      const parsed=parseToast(card);
      if(!parsed?.groupName||!parsed.sender||!parsed.text) return;
      const key=[normalize(parsed.groupName),normalize(parsed.sender),normalize(parsed.text),parsed.timeLabel||card.id||""].join("|");
      const now=Date.now();
      if(!state.toastPrimed) {
        state.toastSeen.set(key,now);
        return;
      }
      if(state.toastSeen.has(key)) return;
      state.toastSeen.set(key,now);
      if(state.toastSeen.size>500){
        for(const [entry,time] of state.toastSeen) if(now-time>300000) state.toastSeen.delete(entry);
        if(state.toastSeen.size>500) state.toastSeen.clear();
      }
      send("TEAMS_LIVE_MESSAGE",{
        groupName:parsed.groupName,
        groupKey:normalize(parsed.groupName),
        sender:parsed.sender,
        text:parsed.text,
        timestamp:now,
        messageTimestamp:parsed.timeLabel||"",
        messageId:"teams-toast:"+key,
        audioRef:parsed.audioRef||"",
        audioUrl:parsed.audioUrl||""
      });
    });
    if(!state.toastPrimed) state.toastPrimed=true;
  }

  function start(){
    if(!/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(location.href)||state.running) return;
    state.running=state.enabled=true;
    state.root=null;
    state.seen.clear();
    scanToastNotifications();
    state.toastObserver=new MutationObserver(scanToastNotifications);
    state.toastObserver.observe(document.documentElement,{childList:true,subtree:true,characterData:true});
    state.toastPoller=setInterval(scanToastNotifications,900);
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
      state.lastSeenKey=baseline.messages.length?baseline.messages[baseline.messages.length-1].key:null;
      state.hasBaseline=baseline.messages.length>0;
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
        state.lastSeenKey=batch.messages.length?batch.messages[batch.messages.length-1].key:null;
        state.hasBaseline=batch.messages.length>0;
        return;
      }

      if(!state.hasBaseline){
        if(batch.messages.length){
          state.seen=new Set(batch.messages.map(message=>message.key));
          state.lastSeenKey=batch.messages[batch.messages.length-1].key;
          state.hasBaseline=true;
          send("TEAMS_LIVE_STATUS",{running:true,text:"Initial Teams history loaded. Existing messages skipped: "+batch.messages.length+"."});
        }
        return;
      }

      const anchorIndex=batch.messages.findIndex(message=>message.key===state.lastSeenKey);
      if(anchorIndex<0){
        // The message list was virtualized/rebuilt. Re-baseline it instead of replaying old history.
        state.seen=new Set(batch.messages.map(message=>message.key));
        state.lastSeenKey=batch.messages.length?batch.messages[batch.messages.length-1].key:null;
        send("TEAMS_LIVE_STATUS",{running:true,text:"Teams message list refreshed. Existing messages skipped; listening from the newest visible message."});
        return;
      }

      const appended=batch.messages.slice(anchorIndex+1);
      const quotedOriginalIds=new Set(appended.map(message=>message.replyToMessageId).filter(Boolean));
      appended.forEach(message=>{
        if(quotedOriginalIds.has(message.key)){
          state.seen.add(message.key);
          return;
        }
        if(state.seen.has(message.key)) return;
        state.seen.add(message.key);
        if(message.own) return;
        send("TEAMS_LIVE_STATUS",{running:true,text:"Detected new message in selected group: "+groupName()+" · "+message.sender+"."});
        send("TEAMS_LIVE_MESSAGE",{
          groupName:groupName(),
          groupKey:groupKey(),
          sender:message.sender,
          text:message.text,
          timestamp:message.timestamp,
          messageTimestamp:message.timestamp||"",
          messageId:message.key,
          replyTo:message.replyTo||"",
          replyToMessageId:message.replyToMessageId||"",
          isReply:!!message.isReply,
          audioRef:message.audioRef||"",
          audioUrl:message.audioUrl||""
        });
      });

      if(batch.messages.length) state.lastSeenKey=batch.messages[batch.messages.length-1].key;
      if(state.seen.size>3000) state.seen=new Set(Array.from(state.seen).slice(-1800));
    }
    state.waitingObserver=new MutationObserver(schedule);
    state.waitingObserver.observe(document.documentElement,{childList:true,subtree:true});
    scan();
    state.poller=setInterval(scan,1200);
    send("TEAMS_LIVE_STATUS",{running:true,text:"Teams live reader ON. Waiting for selected group updates."});
  }

  chrome.runtime.onMessage.addListener((message,sender,sendResponse)=>{
    if(message?.type==="TEAMS_PLAY_AUDIO"){
      const ref=String(message.audioRef||"");
      let audio=ref?Array.from(document.querySelectorAll("audio[data-vs-audio-ref]")).find(node=>node.getAttribute("data-vs-audio-ref")===ref):null;
      if(!audio&&message.audioUrl){
        try{audio=new Audio(message.audioUrl);}catch(_){}
      }
      if(!audio){
        sendResponse({ok:false,error:"Audio player or attachment URL is not available in this Teams tab."});
        return true;
      }
      try{
        state.activeAudio?.pause();
        state.activeAudio=audio;
        audio.muted=false;
        audio.volume=1;
        audio.currentTime=0;
        audio.addEventListener("ended",()=>{
          if(state.activeAudio===audio) state.activeAudio=null;
          chrome.runtime.sendMessage({type:"TEAMS_AUDIO_ENDED",playbackId:message.playbackId},()=>void chrome.runtime.lastError);
        },{once:true});
        audio.addEventListener("error",()=>{
          if(state.activeAudio===audio) state.activeAudio=null;
          chrome.runtime.sendMessage({type:"TEAMS_AUDIO_FAILED",playbackId:message.playbackId,error:"Teams audio playback failed."},()=>void chrome.runtime.lastError);
        },{once:true});
        audio.play().then(()=>sendResponse({ok:true,started:true})).catch(error=>sendResponse({ok:false,error:error?.message||"Browser blocked audio playback."}));
      }catch(error){
        sendResponse({ok:false,error:error?.message||"Could not start audio playback."});
      }
      return true;
    }
    if(message?.type==="TEAMS_STOP_AUDIO"){
      if(state.activeAudio){
        try{state.activeAudio.pause();state.activeAudio.currentTime=0;}catch(_){}
        state.activeAudio=null;
      }
      sendResponse({ok:true});
      return true;
    }
    if(message?.type==="TEAMS_GET_GROUP_INFO_V17"){
      const name=groupName(),key=normalize(name),root=findRoot();
      const batch=root?extract(root.node):{messages:[]};
      const error=!name?"Teams group name could not be detected.":!root?"Open the group chat so its messages are visible.":"";
      sendResponse({ok:!!name&&!!root,groupName:name,groupKey:key,baselineCount:batch.messages.length,error});
      return true;
    }
    if(message?.type==="TEAMS_GET_CHAT_INFO_V16"){
      const info=chatInfo();
      sendResponse(info);
      return true;
    }
    if(message?.type==="TEAMS_READ_LAST_FIVE_V16"){
      const name=groupName(),key=normalize(name),root=findRoot();
      if(!name||!root||key!==message.groupKey){
        sendResponse({ok:false,error:"The current Teams tab is no longer on this selected group. Open that group and try again."});
        return true;
      }
      const batch=extract(root.node);
      let messages=batch.messages.filter(message=>!message.own);
      const quotedOriginalIds=new Set(messages.map(message=>message.replyToMessageId).filter(Boolean));
      messages=messages.filter(message=>!quotedOriginalIds.has(message.key)).slice(-5);
      sendResponse({ok:true,groupName:name,messages});
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
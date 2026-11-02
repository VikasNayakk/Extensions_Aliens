const $=id=>document.getElementById(id);
let pairs=[],selectedId=null,voices=[];
const getActiveTab=async()=>{const [tab]=await chrome.tabs.query({active:true,currentWindow:true});return tab;};
async function loadConversation(){
 $("status").textContent="Reading the active ChatGPT conversation…";
 try{
  const tab=await getActiveTab();
  if(!tab?.id||!/^https:\/\/(chatgpt\.com|chat\.openai\.com)\//.test(tab.url||""))throw new Error("Open a ChatGPT conversation in the active tab first.");
  const [{result}]=await chrome.scripting.executeScript({target:{tabId:tab.id},func:()=>{
   const nodes=[...document.querySelectorAll('[data-message-author-role]')];
   const out=[];let current=null;
   const textOf=n=>(n.querySelector('.markdown')||n.querySelector('[data-message-content]')||n).innerText?.trim()||"";
   for(const n of nodes){
    const role=n.getAttribute("data-message-author-role");
    const text=textOf(n);if(!text)continue;
    if(role==="user"){current={id:out.length+"-"+text.slice(0,24),prompt:text,reply:""};out.push(current);}
    else if(role==="assistant"&&current){current.reply=text;}
   }
   return {title:document.title,messages:out};
  }});
  pairs=(result?.messages||[]).filter(x=>x.prompt);
  renderList();
  if(pairs.length){selectMessage(pairs[pairs.length-1].id);$("status").textContent="Loaded "+pairs.length+" of your messages. Latest prompt selected.";}
  else $("status").textContent="No user messages found. Try refreshing the ChatGPT page.";
 }catch(e){$("status").textContent=e.message||"Could not read this conversation.";}
}
function renderList(){
 const q=$("search").value.trim().toLowerCase();
 const visible=pairs.filter(p=>p.prompt.toLowerCase().includes(q));
 $("count").textContent=visible.length+" / "+pairs.length;
 $("messages").replaceChildren();
 visible.forEach((p,i)=>{
  const b=document.createElement("button");b.className="message-item"+(p.id===selectedId?" active":"");b.type="button";
  const title=document.createElement("strong");title.textContent=(pairs.indexOf(p)+1)+". "+p.prompt.replace(/\s+/g," ").slice(0,115);
  const meta=document.createElement("small");meta.textContent=p.reply?"Reply available":"Reply not yet available";
  b.append(title,meta);b.addEventListener("click",()=>selectMessage(p.id));$("messages").append(b);
 });
}
function selectMessage(id){
 selectedId=id;const p=pairs.find(x=>x.id===id);if(!p)return;
 $("selected").hidden=false;$("selectedTitle").textContent="Reply to message "+(pairs.indexOf(p)+1);
 $("promptText").textContent=p.prompt;$("replyText").textContent=p.reply||"ChatGPT has not replied to this message yet.";
 $("speak").disabled=!p.reply;renderList();
}
$("load").addEventListener("click",loadConversation);$("refresh").addEventListener("click",loadConversation);
$("search").addEventListener("input",renderList);
$("settings").addEventListener("click",()=>chrome.runtime.openOptionsPage());
$("copy").addEventListener("click",async()=>{const p=pairs.find(x=>x.id===selectedId);if(p)await navigator.clipboard.writeText(p.reply);});
$("stop").addEventListener("click",()=>speechSynthesis.cancel());
$("rate").addEventListener("input",()=>{$("rateValue").textContent=Number($("rate").value).toFixed(1)+"×";});
function loadVoices(){voices=speechSynthesis.getVoices();const select=$("voice");const previous=select.value;select.replaceChildren(new Option("Default voice","any"),new Option("Female voice","female"),new Option("Male voice","male"));voices.forEach((v,i)=>{const o=new Option(v.name+" ("+v.lang+")",String(i));select.append(o);});if([...select.options].some(o=>o.value===previous))select.value=previous;}
speechSynthesis.addEventListener("voiceschanged",loadVoices);loadVoices();
$("speak").addEventListener("click",()=>{
 const p=pairs.find(x=>x.id===selectedId);if(!p?.reply)return;speechSynthesis.cancel();
 const u=new SpeechSynthesisUtterance(p.reply);u.rate=Number($("rate").value);
 const choice=$("voice").value;
 if(choice==="female"||choice==="male"){const re=choice==="female"?/female|woman|zira|samantha| victoria|karen|aria|jenny|susan|hazel/i:/male|man|david|mark|guy| george|james|daniel/i;u.voice=voices.find(v=>re.test(v.name))||null;}
 else if(choice!=="any")u.voice=voices[Number(choice)]||null;
 speechSynthesis.speak(u);
});
loadConversation();
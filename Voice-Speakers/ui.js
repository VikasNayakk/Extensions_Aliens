const $ = id => document.getElementById(id);
let pairs = [];
let selectedId = null;
let voices = [];
let speechRun = 0;

function setStatus(message) {
  $("status").textContent = message;
}

let teamsMonitorTabId = null;
let teamsDetectedCount = 0;
let teamsSpeechQueue = [];
let teamsSpeechActive = false;

function teamsMonitorControl(action) {
  const stateKey = "__voiceSpeakersTeamsLiveReaderV13";
  const notify = (type, payload) => {
    try { chrome.runtime.sendMessage({ type, ...payload }, () => void chrome.runtime.lastError); } catch (_) {}
  };

  if (action === "stop") {
    const old = window[stateKey];
    if (old) {
      old.running = false;
      if (old.observer) old.observer.disconnect();
      if (old.timer) clearTimeout(old.timer);
      if (old.poller) clearInterval(old.poller);
      delete window[stateKey];
    }
    notify("TEAMS_LIVE_STATUS", { running: false, text: "Live reading stopped." });
    return { ok: true, stopped: true };
  }

  if (!/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(location.href)) {
    return { ok: false, error: "Open Microsoft Teams first." };
  }

  const existing = window[stateKey];
  if (existing && existing.running) {
    notify("TEAMS_LIVE_STATUS", { running: true, text: "Live reading is already active in this chat." });
    return { ok: true, alreadyRunning: true };
  }

  const rootSelectors = [
    "#chat-pane-list",
    '[data-tid="chat-pane-message-list"]',
    '[data-tid="thread-body-scrollable-content"]',
    '[data-tid="scrollable-thread-body"]',
    'div[role="log"][aria-label*="messages" i]',
    'div[role="list"][aria-label*="messages" i]',
    '[data-tid="message-pane"]'
  ];
  const bodySelector = '[id^="message-body-"]';
  const authorSelector = '[data-tid="message-author-name"]';
  const findRoot = () => rootSelectors.map(selector => document.querySelector(selector)).find(Boolean) || null;
  const clean = value => String(value || "").replace(/\u200b/g, "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();

  function findItems(root) {
    if (!root) return [];
    const direct = Array.from(root.children || []).filter(node => node.querySelector(bodySelector));
    if (direct.length) return direct;

    const selectors = [
      '[data-tid="message-list-item"]',
      '[data-tid="message-group-container"]',
      'div[role="listitem"]'
    ];
    for (const selector of selectors) {
      const items = Array.from(root.querySelectorAll(selector)).filter(node => node.querySelector(bodySelector));
      if (items.length) return items.filter(node => !items.some(other => other !== node && other.contains(node)));
    }

    return Array.from(root.querySelectorAll(bodySelector)).map(body =>
      body.closest('[data-tid="message-list-item"], [data-tid="message-group-container"], div[role="listitem"]') || body.parentElement
    ).filter((node, index, all) => node && all.indexOf(node) === index);
  }

  function closestAuthor(item) {
    let node = item;
    for (let depth = 0; node && depth < 5; depth++, node = node.parentElement) {
      const author = node.querySelector(authorSelector);
      if (author && clean(author.innerText || author.textContent)) return clean(author.innerText || author.textContent);
    }
    return "";
  }

  function isOutgoing(item, author) {
    if (/^(you|me)$/i.test(author)) return true;
    let node = item;
    for (let depth = 0; node && depth < 7; depth++, node = node.parentElement) {
      const classes = typeof node.className === "string" ? node.className : "";
      const markers = [
        classes,
        node.getAttribute("data-tid"),
        node.getAttribute("data-testid"),
        node.getAttribute("data-message-direction"),
        node.getAttribute("data-is-own-message"),
        node.getAttribute("aria-label")
      ].filter(Boolean).join(" ").toLowerCase();
      if (/outgoing|message-out|own-message|self-message|sent-by-me|from-me|data-is-own-message true|you sent|you said/.test(markers)) return true;
    }
    return false;
  }

  function redactSensitive(value) {
    return clean(value)
      .replace(/\b(password|passcode|passwd|pwd|otp|one[- ]time code|api[ _-]?key|access[ _-]?token|secret|2fa)\b\s*(?:is\s*)?[:=]?\s*([^\s,;]+)/gi, "$1 [hidden]")
      .replace(/\bBearer\s+[A-Za-z0-9._~+\/-]+=*/gi, "Bearer [hidden]")
      .replace(/\b(?:sk-[A-Za-z0-9_-]{16,}|gh[pousr]_[A-Za-z0-9]{20,})\b/g, "[access token hidden]");
  }

  function readEntries(root) {
    const items = findItems(root);
    const entries = [];
    let lastAuthor = "";
    items.forEach((item, itemIndex) => {
      const bodies = Array.from(item.matches?.(bodySelector) ? [item] : item.querySelectorAll(bodySelector));
      const author = closestAuthor(item) || lastAuthor;
      if (author) lastAuthor = author;
      bodies.forEach(body => {
        const timeNode = item.querySelector('[id^="timestamp-"]') || item.querySelector("time[datetime]");
        const content = body.querySelector('[id^="content-"]') || body;
        const clone = content.cloneNode(true);
        clone.querySelectorAll("button,[role='button'],svg,script,style,[aria-hidden='true'],[data-tid*='reaction'],[data-testid*='reaction']").forEach(node => node.remove());
        const rawText = clean(clone.innerText || clone.textContent);
        const text = redactSensitive(rawText);
        if (!text || !author) return;
        const timestamp = timeNode
          ? clean(timeNode.getAttribute("datetime") || timeNode.innerText || timeNode.textContent)
          : "";
        const id = body.id || "";
        const key = id || [author, timestamp, rawText].join("|");
        entries.push({
          key,
          sender: author,
          text,
          timestamp,
          outgoing: isOutgoing(item, author),
          itemIndex
        });
      });
    });
    return entries;
  }

  function conversationKey(root) {
    const header = document.querySelector('[data-tid="entity-header"]') ||
      document.querySelector('[data-tid="thread-header"]') ||
      document.querySelector('[data-tid="chat-header"]');
    const title = clean(header && (header.innerText || header.textContent));
    return location.pathname + "|" + (title || (root && (root.getAttribute("data-tid") || root.id)) || "teams-chat");
  }

  const root = findRoot();
  if (!root) return { ok: false, error: "Teams is open, but I couldn't find the chat message list. Open an individual chat and try again." };
  const initialEntries = readEntries(root);
  if (!initialEntries.length) return { ok: false, error: "I found the Teams chat area, but no message items. Open a chat with messages and try again." };

  const state = {
    running: true,
    root,
    chatKey: conversationKey(root),
    seen: new Set(initialEntries.map(entry => entry.key)),
    observer: null,
    timer: null,
    poller: null
  };
  window[stateKey] = state;

  function scan() {
    if (!state.running) return;
    const currentRoot = findRoot();
    if (!currentRoot) return;
    const currentChatKey = conversationKey(currentRoot);
    const entries = readEntries(currentRoot);

    if (currentRoot !== state.root || currentChatKey !== state.chatKey) {
      state.root = currentRoot;
      state.chatKey = currentChatKey;
      state.seen = new Set(entries.map(entry => entry.key));
      notify("TEAMS_LIVE_STATUS", { running: true, text: "Chat changed. Previous messages skipped; listening for new messages here." });
      return;
    }

    entries.forEach((entry, index) => {
      if (state.seen.has(entry.key)) return;
      state.seen.add(entry.key);

      // Messages added while loading older history usually appear at the top.
      // Only announce newly appended items near the bottom of the active chat.
      const isNearLatest = index >= Math.max(0, entries.length - 5);
      if (!isNearLatest || entry.outgoing) return;

      notify("TEAMS_LIVE_MESSAGE", {
        sender: entry.sender,
        text: entry.text,
        timestamp: entry.timestamp,
        messageId: entry.key
      });
    });

    if (state.seen.size > 2500) state.seen = new Set(Array.from(state.seen).slice(-1500));
  }

  const scheduleScan = () => {
    if (!state.running || state.timer) return;
    state.timer = setTimeout(() => {
      state.timer = null;
      scan();
    }, 250);
  };

  state.observer = new MutationObserver(scheduleScan);
  state.observer.observe(document.body, { childList: true, subtree: true, characterData: true });
  state.poller = setInterval(scan, 1500);
  notify("TEAMS_LIVE_STATUS", {
    running: true,
    text: "Live reading is on. Listening for new incoming messages in this chat; existing messages were skipped."
  });
  return { ok: true, baselineCount: initialEntries.length };
}

function addTeamsFeedItem(sender, message) {
  const feed = $("teamsFeed");
  if (!feed) return;
  const empty = feed.querySelector(".empty");
  if (empty) empty.remove();
  const item = document.createElement("div");
  item.className = "teams-feed-item";
  const name = document.createElement("strong");
  name.textContent = sender;
  const body = document.createElement("span");
  body.textContent = message;
  item.append(name, body);
  feed.prepend(item);
  while (feed.children.length > 8) feed.lastElementChild.remove();
}

function speakTeamsQueue() {
  if (teamsSpeechActive || !teamsSpeechQueue.length || !("speechSynthesis" in window)) return;
  const entry = teamsSpeechQueue.shift();
  const chunks = splitForSpeech(entry.text, 200);
  let first = true;
  teamsSpeechActive = true;

  const readNext = () => {
    if (!chunks.length) {
      teamsSpeechActive = false;
      speakTeamsQueue();
      return;
    }
    let part = chunks.shift();
    if (first) {
      part = entry.sender + " ne kaha: " + part;
      first = false;
    }
    const utterance = new SpeechSynthesisUtterance(part);
    utterance.rate = Number($("rate")?.value) || 0.8;
    const choice = $("voice")?.value || "auto";
    let voice = null;
    if (choice.startsWith("v")) voice = voices[Number(choice.slice(1))] || null;
    else if (choice !== "auto") {
      const pattern = choice === "female"
        ? /female|zira|samantha|victoria|karen|aria|jenny|susan|hazel|ava/i
        : /male|david|mark|guy|george|james|daniel/i;
      voice = voices.find(candidate => pattern.test(candidate.name)) || null;
    }
    if (voice) {
      utterance.voice = voice;
      utterance.lang = voice.lang || "hi-IN";
    } else {
      utterance.lang = "hi-IN";
    }
    utterance.onend = readNext;
    utterance.onerror = () => {
      teamsSpeechActive = false;
      speakTeamsQueue();
    };
    window.speechSynthesis.speak(utterance);
  };

  window.speechSynthesis.resume();
  readNext();
}

async function startTeamsLiveReading() {
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  if (!tab || !tab.id || !/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(tab.url || "")) {
    $("teamsStatus").textContent = "Open Microsoft Teams in the active tab, then start live reading.";
    return;
  }
  $("teamsStart").disabled = true;
  $("teamsStatus").textContent = "Connecting to the open Teams chat…";
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: teamsMonitorControl,
      args: ["start"]
    });
    const result = results && results[0] && results[0].result;
    if (!result || !result.ok) throw new Error(result?.error || "Could not start the Teams monitor.");
    teamsMonitorTabId = tab.id;
    await chrome.storage.local.set({ teamsMonitorTabId });
    $("teamsStop").disabled = false;
    $("teamsStatus").textContent = "Live reading is on. New incoming messages will be announced with the sender's name.";
  } catch (error) {
    $("teamsStatus").textContent = error?.message || "Could not start live reading.";
  } finally {
    $("teamsStart").disabled = false;
  }
}

async function stopTeamsLiveReading() {
  try {
    const tabId = teamsMonitorTabId || (await chrome.tabs.query({ active: true, lastFocusedWindow: true }))[0]?.id;
    if (tabId) await chrome.scripting.executeScript({ target: { tabId }, func: teamsMonitorControl, args: ["stop"] });
  } catch (_) {}
  teamsMonitorTabId = null;
  await chrome.storage.local.remove("teamsMonitorTabId");
  teamsSpeechQueue = [];
  teamsSpeechActive = false;
  if ("speechSynthesis" in window) window.speechSynthesis.cancel();
  $("teamsStart").disabled = false;
  $("teamsStop").disabled = true;
  $("teamsStatus").textContent = "Live reading stopped.";
}

async function initializeActiveTab() {
  const settings = await chrome.storage.local.get(["teamsMonitorTabId"]);
  teamsMonitorTabId = settings.teamsMonitorTabId || null;
  if (teamsMonitorTabId) {
    $("teamsStop").disabled = false;
    $("teamsStatus").textContent = "A Teams monitor may still be running. Keep this panel open to hear new messages.";
  }
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  if (/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(tab?.url || "")) {
    $("teamsStatus").textContent = "Teams detected. Start live reading to announce new incoming messages.";
  } else {
    loadConversation();
  }
}

chrome.runtime.onMessage.addListener(message => {
  if (message?.type === "TEAMS_LIVE_STATUS") {
    if ($("teamsStatus")) {
      $("teamsStatus").textContent = message.text || (message.running ? "Live reading is on." : "Live reading stopped.");
      $("teamsStop").disabled = !message.running;
      $("teamsStart").disabled = !!message.running;
    }
  }
  if (message?.type === "TEAMS_LIVE_MESSAGE") {
    const sender = String(message.sender || "Someone").trim();
    const text = String(message.text || "").trim();
    if (!text) return;
    teamsDetectedCount += 1;
    $("teamsReadCount").textContent = String(teamsDetectedCount);
    addTeamsFeedItem(sender, text);
    teamsSpeechQueue.push({ sender, text });
    speakTeamsQueue();
  }
});

async function loadConversation() {
  $("status").textContent = "Detecting messages in the open ChatGPT conversation…";
  $("load").disabled = true;
  $("refresh").disabled = true;
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (!tab || !tab.id || !/^https:\/\/(chatgpt\.com|chat\.openai\.com)\//.test(tab.url || "")) {
      throw new Error("Open the ChatGPT conversation tab, then load the conversation.");
    }

    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        const clean = value => String(value || "")
          .replace(/\u200b/g, "")
          .replace(/\u00a0/g, " ")
          .replace(/\r/g, "")
          .trim();

        const visible = element => !!element && (
          element.getClientRects().length > 0 ||
          element.offsetWidth > 0 ||
          element.offsetHeight > 0
        );

        const extractText = (element, role) => {
          // Current ChatGPT markup places the role on a visually-hidden heading
          // and the actual message body beside it in the same block.
          const scope = element.closest(".block-BQZwFn") || element;
          let content = scope;
          if (role === "user") {
            content = scope.querySelector("[data-user-message-bubble]") ||
              scope.querySelector(".whitespace-pre-wrap") ||
              scope.querySelector("[data-message-content]") ||
              scope;
          } else {
            content = scope.querySelector('[data-markdown-text-style="assistant-message"]') ||
              scope.querySelector(".markdown") ||
              scope.querySelector("[data-message-content]") ||
              scope;
          }

          const clone = content.cloneNode(true);
          clone.querySelectorAll(
            "button, [role='button'], svg, script, style, nav, [aria-hidden='true'], [data-testid*='action'], [data-testid*='toolbar']"
          ).forEach(node => node.remove());

          return clean(clone.innerText || clone.textContent);
        };

        const modernRoleNodes = Array.from(document.querySelectorAll(
          'main [data-conversation-role="user"], main [data-conversation-role="assistant"], [data-conversation-role="user"], [data-conversation-role="assistant"]'
        )).filter(visible);

        const rawRoleNodes = Array.from(document.querySelectorAll(
          'main [data-message-author-role="user"], main [data-message-author-role="assistant"], [data-message-author-role="user"], [data-message-author-role="assistant"]'
        )).filter(visible);

        let messages = [];
        const normalize = value => clean(value).replace(/\s+/g, " ").toLowerCase();

        if (modernRoleNodes.length || rawRoleNodes.length) {
          // Prefer ChatGPT's current semantic role markers. Retain the older
          // markers for compatibility with earlier ChatGPT layouts.
          const roleAttribute = modernRoleNodes.length
            ? "data-conversation-role"
            : "data-message-author-role";
          const roleNodes = (modernRoleNodes.length ? modernRoleNodes : rawRoleNodes).filter(node => {
            const role = node.getAttribute(roleAttribute);
            return !node.parentElement?.closest("[" + roleAttribute + '="' + role + '"]');
          });

          const seenInTurn = new Set();
          roleNodes.forEach((node, index) => {
            const role = node.getAttribute(roleAttribute);
            const turn = node.closest('[data-testid^="conversation-turn-"]') ||
              node.closest(".block-BQZwFn") || node.closest("article") || node.parentElement;
            const text = extractText(node, role);
            if (!text) return;

            // Suppress duplicate DOM renderings only within the same turn.
            // Identical prompts in different turns are retained.
            const turnKey = turn && turn.getAttribute("data-testid")
              ? turn.getAttribute("data-testid")
              : "turn-index-" + index;
            const dedupeKey = role + "::" + turnKey + "::" + normalize(text);
            if (seenInTurn.has(dedupeKey)) return;
            seenInTurn.add(dedupeKey);
            messages.push({ role, text, order: index });
          });

          messages.sort((a, b) => a.order - b.order);
        } else {
          // Fallback: inspect individual conversation turns, then merge by DOM order.
          const turns = Array.from(document.querySelectorAll(
            'main [data-testid^="conversation-turn-"], main article, [data-testid^="conversation-turn-"], article'
          )).filter(visible);

          const uniqueTurns = turns.filter((turn, index) =>
            !turns.slice(0, index).some(parent => parent.contains(turn))
          );

          uniqueTurns.forEach((turn, index) => {
            const userNodes = Array.from(turn.querySelectorAll(
              '[data-testid^="user-message"], [data-message-author-role="user"]'
            ));
            const assistantNodes = Array.from(turn.querySelectorAll(
              '[data-testid^="assistant-message"], [data-message-author-role="assistant"], .markdown'
            ));

            const local = userNodes.map(node => ({ node, role: "user" }))
              .concat(assistantNodes.map(node => ({ node, role: "assistant" })))
              .filter(item => !item.node.parentElement?.closest(
                item.role === "user"
                  ? '[data-testid^="user-message"], [data-message-author-role="user"]'
                  : '[data-testid^="assistant-message"], [data-message-author-role="assistant"]'
              ))
              .sort((a, b) => {
                const position = a.node.compareDocumentPosition(b.node);
                return position & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
              });

            local.forEach(item => {
              const text = extractText(item.node, item.role);
              if (text) messages.push({ role: item.role, text, order: index });
            });
          });
        }

        // Pair each user message with the assistant message(s) that follow it,
        // stopping at the next user message. Equal prompt text in separate turns
        // is intentionally preserved as two distinct prompts.
        const output = [];
        let current = null;
        for (const message of messages) {
          if (message.role === "user") {
            if (current) output.push(current);
            current = { prompt: message.text, replyParts: [] };
          } else if (message.role === "assistant" && current) {
            if (message.text && !current.replyParts.includes(message.text)) {
              current.replyParts.push(message.text);
            }
          }
        }
        if (current) output.push(current);

        return output.map(item => ({
          prompt: clean(item.prompt),
          reply: clean(item.replyParts.join("\n\n"))
        })).filter(item => item.prompt);
      }
    });

    const extracted = results && results[0] ? results[0].result : [];
    pairs = (Array.isArray(extracted) ? extracted : []).map((item, index) => ({
      prompt: String(item.prompt || "").trim(),
      reply: String(item.reply || "").trim(),
      id: String(index),
      index: index + 1
    })).filter(item => item.prompt);

    $("totalCount").textContent = String(pairs.length);
    render();

    if (pairs.length) {
      const previous = pairs.find(item => item.id === selectedId);
      const latest = pairs[pairs.length - 1];
      select(previous ? previous.id : latest.id);
      setStatus(
        "Found " + pairs.length + " distinct prompt message(s). " +
        pairs.filter(item => item.reply).length + " have a detected ChatGPT reply."
      );
    } else {
      selectedId = null;
      $("selected").hidden = true;
      setStatus("No prompts detected. Make sure the conversation is open and its messages are loaded, then refresh.");
    }
    return true;
  } catch (error) {
    pairs = [];
    selectedId = null;
    $("totalCount").textContent = "0";
    $("selected").hidden = true;
    render();
    setStatus(error && error.message ? error.message : "Could not read this conversation.");
    return false;
  } finally {
    $("load").disabled = false;
    $("refresh").disabled = false;
  }
}

function render() {
  const query = $("search").value.toLowerCase().trim();
  const shown = pairs.filter(item => item.prompt.toLowerCase().includes(query));
  $("count").textContent = String(shown.length);
  $("clearSearch").hidden = !query;
  $("messages").replaceChildren();

  if (!shown.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = pairs.length ? "No matching prompts." : "Load a conversation to see prompts.";
    $("messages").append(empty);
    return;
  }

  shown.forEach(item => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "message-item" + (item.id === selectedId ? " active" : "");

    const top = document.createElement("div");
    top.className = "message-top";

    const number = document.createElement("span");
    number.className = "message-number";
    number.textContent = "PROMPT " + item.index;

    const badge = document.createElement("span");
    badge.className = "reply-badge " + (item.reply ? "ready" : "pending");
    badge.textContent = item.reply ? "Reply ready" : "Reply not detected";

    top.append(number, badge);

    const prompt = document.createElement("strong");
    prompt.textContent = item.prompt.replace(/\s+/g, " ").slice(0, 160);
    button.append(top, prompt);
    button.addEventListener("click", () => select(item.id));
    $("messages").append(button);
  });
}

function select(id) {
  selectedId = id;
  const item = pairs.find(entry => entry.id === id);
  if (!item) return;

  $("selected").hidden = false;
  $("selectedTitle").textContent = "Prompt " + item.index;
  $("promptText").textContent = item.prompt;
  $("replyText").textContent = item.reply || "No assistant reply was found after this prompt.";
  $("speakSelected").disabled = !item.reply || !("speechSynthesis" in window);
  $("speakLatest").disabled = !pairs.length || !pairs[pairs.length - 1].reply || !("speechSynthesis" in window);
  render();
}

function voiceList() {
  if (!("speechSynthesis" in window)) return;
  voices = window.speechSynthesis.getVoices();
  const selectVoice = $("voice");
  const oldValue = selectVoice.value;
  selectVoice.replaceChildren(
    new Option("System default", "auto"),
    new Option("Female-sounding voice", "female"),
    new Option("Male-sounding voice", "male")
  );
  voices.forEach((voice, index) => {
    selectVoice.add(new Option(voice.name + " (" + voice.lang + ")", "v" + index));
  });
  if (Array.from(selectVoice.options).some(option => option.value === oldValue)) {
    selectVoice.value = oldValue;
  }
}

function splitForSpeech(value, limit) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (!text) return [];
  const sentences = text.match(/[^.!?।]+[.!?।]?/g) || [text];
  const chunks = [];
  let current = "";

  const addLongSentence = sentence => {
    const words = sentence.split(/\s+/);
    let part = "";
    words.forEach(word => {
      if (part && (part + " " + word).length > limit) {
        chunks.push(part);
        part = word;
      } else {
        part = part ? part + " " + word : word;
      }
    });
    if (part) chunks.push(part);
  };

  sentences.forEach(sentence => {
    const part = sentence.trim();
    if (!part) return;
    if (part.length > limit) {
      if (current) {
        chunks.push(current);
        current = "";
      }
      addLongSentence(part);
    } else if (current && (current + " " + part).length > limit) {
      chunks.push(current);
      current = part;
    } else {
      current = current ? current + " " + part : part;
    }
  });

  if (current) chunks.push(current);
  return chunks;
}

function stopSpeaking(updateStatus) {
  speechRun += 1;
  if ("speechSynthesis" in window) window.speechSynthesis.cancel();
  if (updateStatus !== false) setStatus("Reading stopped.");
}

function speak(item) {
  if (!item || !item.reply) {
    setStatus("This prompt does not have a detected ChatGPT reply. Refresh after the reply has finished generating.");
    return;
  }
  if (!("speechSynthesis" in window) || typeof SpeechSynthesisUtterance === "undefined") {
    setStatus("Text-to-speech is not available in this browser.");
    return;
  }

  stopSpeaking(false);
  const run = speechRun;
  const chunks = splitForSpeech(item.reply, 220);
  const rate = Number($("rate").value) || 0.8;
  const choice = $("voice").value;
  let voice = null;

  if (choice.startsWith("v")) {
    voice = voices[Number(choice.slice(1))] || null;
  } else if (choice !== "auto") {
    const pattern = choice === "female"
      ? /female|zira|samantha|victoria|karen|aria|jenny|susan|hazel|ava/i
      : /male|david|mark|guy|george|james|daniel/i;
    voice = voices.find(candidate => pattern.test(candidate.name)) || null;
  }

  setStatus("Reading reply for prompt " + item.index + "…");

  const readNext = () => {
    if (run !== speechRun) return;
    if (!chunks.length) {
      setStatus("Finished reading prompt " + item.index + ".");
      return;
    }

    const utterance = new SpeechSynthesisUtterance(chunks.shift());
    utterance.rate = rate;
    if (voice) utterance.voice = voice;
    if (voice && voice.lang) utterance.lang = voice.lang;
    utterance.onend = readNext;
    utterance.onerror = event => {
      if (run === speechRun) {
        setStatus("Speech playback error: " + (event.error || "unknown error") + ".");
      }
    };
    window.speechSynthesis.speak(utterance);
  };

  window.speechSynthesis.resume();
  readNext();
}

async function loadSettings() {
  const settings = await chrome.storage.local.get(["apiKey", "model", "voiceRate", "voiceChoice"]);
  $("apiKey").value = settings.apiKey || "";
  $("model").value = settings.model || "gpt-4.1-mini";
  if (settings.voiceRate) {
    $("rate").value = settings.voiceRate;
    $("rateValue").textContent = Number(settings.voiceRate).toFixed(1) + "×";
  }
  if (settings.voiceChoice) $("voice").value = settings.voiceChoice;
}

$("teamsStart").addEventListener("click", startTeamsLiveReading);
$("teamsStop").addEventListener("click", stopTeamsLiveReading);
$("load").addEventListener("click", loadConversation);
$("refresh").addEventListener("click", loadConversation);
$("search").addEventListener("input", render);
$("clearSearch").addEventListener("click", () => {
  $("search").value = "";
  render();
});
$("settings").addEventListener("click", () => chrome.runtime.openOptionsPage());
$("speakSelected").addEventListener("click", () => speak(pairs.find(item => item.id === selectedId)));
$("speakLatest").addEventListener("click", async () => {
  const loaded = await loadConversation();
  if (loaded && pairs.length) speak(pairs[pairs.length - 1]);
});
$("stop").addEventListener("click", () => stopSpeaking(true));
$("copy").addEventListener("click", async () => {
  const item = pairs.find(entry => entry.id === selectedId);
  if (item && item.reply) {
    await navigator.clipboard.writeText(item.reply);
    $("copy").textContent = "Copied";
    setTimeout(() => { $("copy").textContent = "Copy reply"; }, 1000);
  }
});
$("rate").addEventListener("input", () => {
  $("rateValue").textContent = Number($("rate").value).toFixed(1) + "×";
  chrome.storage.local.set({ voiceRate: Number($("rate").value) });
});
$("voice").addEventListener("change", () => {
  chrome.storage.local.set({ voiceChoice: $("voice").value });
});
$("saveKey").addEventListener("click", async () => {
  const key = $("apiKey").value.trim();
  if (!key) {
    $("apiStatus").textContent = "Paste an API key first.";
    return;
  }
  await chrome.storage.local.set({
    apiKey: key,
    model: $("model").value.trim() || "gpt-4.1-mini"
  });
  $("apiStatus").textContent = "Key saved in extension storage.";
});
$("testKey").addEventListener("click", async () => {
  const key = $("apiKey").value.trim();
  if (key) {
    await chrome.storage.local.set({
      apiKey: key,
      model: $("model").value.trim() || "gpt-4.1-mini"
    });
  }
  $("apiStatus").textContent = "Testing API…";
  try {
    const result = await chrome.runtime.sendMessage({ type: "TEST_OPENAI" });
    $("apiStatus").textContent = result && result.ok
      ? "API connection successful (" + result.model + ")."
      : (result && result.error) || "API test failed.";
  } catch (error) {
    $("apiStatus").textContent = "Test failed: " + error.message;
  }
});

if ("speechSynthesis" in window) {
  window.speechSynthesis.onvoiceschanged = voiceList;
  voiceList();
} else {
  $("speakSelected").disabled = true;
  $("speakLatest").disabled = true;
  setStatus("This browser does not provide speech synthesis.");
}
loadSettings();
initializeActiveTab();

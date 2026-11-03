const $ = id => document.getElementById(id);
let pairs = [];
let selectedId = null;
let voices = [];
let speechRun = 0;

function setStatus(message) {
  $("status").textContent = message;
}

let teamsMonitorTabId = null;
let teamsMonitorFrameId = null;
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
    '[data-tid="thread-body-scrollable-content"]',
    '[data-tid="message-pane-body"]',
    '[data-tid="chat-pane-message-list"]',
    '[data-tid="scrollable-thread-body"]',
    'div[role="log"][aria-label*="messages" i]',
    'div[role="list"][aria-label*="messages" i]',
    '[data-tid="thread-view"]',
    '[data-tid="message-pane"]',
    "#chat-pane-list"
  ];
  const messageSelectors = [
    '[id^="message-body-"]',
    '[data-tid="message-content"]',
    '[data-tid="message-body"]',
    '[data-tid="bubble-text"]',
    '[data-tid^="message-content-"]',
    '[data-tid^="messageBody"]',
    ".ts-message-content",
    ".message-body",
    ".ui-chat__messagecontent",
    '[role="document"]'
  ];
  const itemSelectors = [
    '[data-tid="message-group-container"]',
    '[data-tid="message-list-item"]',
    'div[role="listitem"]',
    'div[data-tid^="message-"]'
  ];
  const authorSelectors = [
    '[data-tid="message-author-name"]',
    '[data-tid="author-name"]',
    '[id^="message-author-"]',
    'span[data-tid="message-group-author"]',
    '[data-tid="message-group-header"]'
  ];
  const timestampSelectors = [
    '[id^="timestamp-"]',
    '[data-tid="message-time"]',
    '[data-tid="message-group-time"]',
    '[data-tid="message-timestamp"]',
    "time[datetime]",
    "time"
  ];
  const clean = value => String(value || "").replace(/\u200b/g, "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const visible = element => !!element && element.getClientRects().length > 0;
  const findRootMatch = () => {
    for (const selector of rootSelectors) {
      const root = document.querySelector(selector);
      if (root && visible(root)) return { root, selector };
    }
    return { root: null, selector: "" };
  };
  const findRoot = () => findRootMatch().root;

  function findMessageNodes(container) {
    if (!container) return [];
    for (const selector of messageSelectors) {
      const nodes = [];
      if (container.matches?.(selector) && visible(container)) nodes.push(container);
      nodes.push(...Array.from(container.querySelectorAll(selector)).filter(visible));
      if (nodes.length) {
        return nodes.filter((node, index) =>
          !nodes.some((other, otherIndex) => otherIndex !== index && other.contains(node))
        );
      }
    }
    return [];
  }

  function findItems(root) {
    if (!root) return [];
    for (const selector of itemSelectors) {
      const candidates = Array.from(root.querySelectorAll(selector))
        .filter(node => visible(node) && findMessageNodes(node).length > 0);
      if (candidates.length) {
        return candidates.filter(node =>
          !candidates.some(other => other !== node && other.contains(node))
        );
      }
    }

    const direct = Array.from(root.children || [])
      .filter(node => visible(node) && findMessageNodes(node).length > 0);
    if (direct.length) return direct;

    return findMessageNodes(root).map(body =>
      body.closest(itemSelectors.join(",")) || body.parentElement
    ).filter((node, index, all) => node && all.indexOf(node) === index);
  }

  function findMetadata(startNode, selectors) {
    let node = startNode;
    for (let depth = 0; node && depth < 6; depth++, node = node.parentElement) {
      for (const selector of selectors) {
        const found = node.matches?.(selector) ? node : node.querySelector(selector);
        if (!found) continue;
        const value = clean(found.getAttribute("aria-label") || found.getAttribute("title") || found.innerText || found.textContent);
        if (value) return value;
      }
    }
    return "";
  }

  function currentUserName() {
    const candidates = Array.from(document.querySelectorAll("[data-tid], [role='treeitem'], [role='listitem']"));
    for (const node of candidates) {
      if (!visible(node)) continue;
      const text = clean(node.innerText || node.textContent);
      const match = text.match(/^(.{2,80}?)\s*\(You\)$/i);
      if (match) return clean(match[1]);
    }
    return "";
  }
  const selfDisplayName = currentUserName();

  function isOutgoing(item, body, author) {
    if (/^(you|me)$/i.test(author)) return true;
    if (selfDisplayName && author.toLowerCase() === selfDisplayName.toLowerCase()) return true;
    let node = body || item;
    for (let depth = 0; node && depth < 8; depth++, node = node.parentElement) {
      const classes = typeof node.className === "string" ? node.className : "";
      const markers = [
        classes,
        node.getAttribute("data-tid"),
        node.getAttribute("data-testid"),
        node.getAttribute("data-message-direction"),
        node.getAttribute("data-is-own-message"),
        node.getAttribute("aria-label"),
        node.getAttribute("data-message-author")
      ].filter(Boolean).join(" ").toLowerCase();
      if (/outgoing|message-out|own-message|self-message|sent-by-me|from-me|data-is-own-message true|you sent|you said/.test(markers)) return true;
      if (node === item) break;
    }
    return false;
  }

  function redactSensitive(value) {
    return clean(value)
      .replace(/\b(password|passcode|passwd|pwd|otp|one[- ]time code|api[ _-]?key|access[ _-]?token|secret|2fa)\b\s*(?:is\s*)?[:=]?\s*([^\s,;]+)/gi, "$1 [hidden]")
      .replace(/\bBearer\s+[A-Za-z0-9._~+\/-]+=*/gi, "Bearer [hidden]")
      .replace(/\b(?:sk-[A-Za-z0-9_-]{16,}|gh[pousr]_[A-Za-z0-9]{20,})\b/g, "[access token hidden]")
      .replace(/\b(?=[A-Za-z0-9!@#$%^&*_-]{12,}\b)(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[!@#$%^&*_-])[A-Za-z0-9!@#$%^&*_-]+\b/g, "[possible password hidden]");
  }

  function readEntries(root) {
    const items = findItems(root);
    const entries = [];
    let lastAuthor = "";
    items.forEach((item, itemIndex) => {
      const bodies = findMessageNodes(item);
      bodies.forEach(body => {
        const author = findMetadata(body, authorSelectors) || findMetadata(item, authorSelectors) || lastAuthor;
        if (author) lastAuthor = author;

        let content = body;
        const nestedContent = body.querySelector('[id^="content-"]');
        if (nestedContent) content = nestedContent;
        const clone = content.cloneNode(true);
        clone.querySelectorAll(
          "button,[role='button'],svg,script,style,nav,[aria-hidden='true'],[data-tid*='reaction'],[data-testid*='reaction']"
        ).forEach(node => node.remove());

        const rawText = clean(clone.innerText || clone.textContent);
        const text = redactSensitive(rawText);
        if (!text || !author) return;

        const timestamp = findMetadata(body, timestampSelectors) || findMetadata(item, timestampSelectors);
        const messageId = body.id ||
          body.getAttribute("data-message-id") ||
          item.getAttribute("data-message-id") ||
          "";
        const key = messageId || [author, timestamp, rawText].join("|");

        entries.push({
          key,
          sender: author.replace(/^Message by\s+/i, "").replace(/\s+by\s+.+$/i, ""),
          text,
          timestamp,
          outgoing: isOutgoing(item, body, author),
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

  const rootMatch = findRootMatch();
  const root = rootMatch.root;
  if (!root) return { ok: false, error: "Teams chat message area not found in this frame. Checked: " + rootSelectors.join(", ") };
  const initialItems = findItems(root);
  const initialEntries = readEntries(root);
  if (!initialEntries.length) {
    return {
      ok: false,
      error: "Teams message pane found (" + rootMatch.selector + "), but no readable message records. Message wrappers: " + initialItems.length + ". Try opening the chat and scrolling until messages are visible."
    };
  }

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
  return {
    ok: true,
    baselineCount: initialEntries.length,
    messageWrappers: initialItems.length,
    rootSelector: rootMatch.selector,
    frameUrl: location.href,
    selfNameDetected: !!selfDisplayName
  };
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
      target: { tabId: tab.id, allFrames: true },
      func: teamsMonitorControl,
      args: ["start"]
    });
    const successfulFrame = (results || []).find(entry => entry.result && entry.result.ok);
    if (!successfulFrame) {
      const errors = (results || []).map(entry => entry.result && entry.result.error).filter(Boolean);
      throw new Error(errors[0] || "No Teams frame contained a readable chat. Open a chat and try again.");
    }
    const result = successfulFrame.result;
    teamsMonitorTabId = tab.id;
    teamsMonitorFrameId = successfulFrame.frameId;
    await chrome.storage.local.set({ teamsMonitorTabId, teamsMonitorFrameId });
    $("teamsStop").disabled = false;
    $("teamsStatus").textContent =
      "Listening for new incoming messages. Baseline skipped: " + result.baselineCount +
      " messages; wrapper count: " + result.messageWrappers +
      "; frame: " + successfulFrame.frameId + "; source: " + result.rootSelector +
      (result.selfNameDetected ? "; own messages filtered." : "; Teams did not expose the signed-in name, so verify your own messages are not announced.");
  } catch (error) {
    $("teamsStatus").textContent = error?.message || "Could not start live reading.";
  } finally {
    $("teamsStart").disabled = false;
  }
}

async function stopTeamsLiveReading() {
  try {
    const tabId = teamsMonitorTabId || (await chrome.tabs.query({ active: true, lastFocusedWindow: true }))[0]?.id;
    if (tabId) {
      const target = Number.isInteger(teamsMonitorFrameId)
        ? { tabId, frameIds: [teamsMonitorFrameId] }
        : { tabId, allFrames: true };
      await chrome.scripting.executeScript({ target, func: teamsMonitorControl, args: ["stop"] });
    }
  } catch (_) {}
  teamsMonitorTabId = null;
  teamsMonitorFrameId = null;
  await chrome.storage.local.remove(["teamsMonitorTabId", "teamsMonitorFrameId"]);
  teamsSpeechQueue = [];
  teamsSpeechActive = false;
  if ("speechSynthesis" in window) window.speechSynthesis.cancel();
  $("teamsStart").disabled = false;
  $("teamsStop").disabled = true;
  $("teamsStatus").textContent = "Live reading stopped.";
}

async function initializeActiveTab() {
  const settings = await chrome.storage.local.get(["teamsMonitorTabId", "teamsMonitorFrameId"]);
  teamsMonitorTabId = settings.teamsMonitorTabId || null;
  teamsMonitorFrameId = Number.isInteger(settings.teamsMonitorFrameId) ? settings.teamsMonitorFrameId : null;
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

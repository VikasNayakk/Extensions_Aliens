(() => {
  const STATE_KEY = "__aliensVoiceSpeakersTeamsContentV14";
  if (window[STATE_KEY]) return;

  const state = window[STATE_KEY] = {
    enabled: false,
    running: false,
    root: null,
    rootSelector: "",
    chatKey: "",
    seen: new Set(),
    observer: null,
    waitingObserver: null,
    poller: null,
    timer: null,
    suppressUntil: 0
  };

  const send = (type, payload = {}) => {
    try { chrome.runtime.sendMessage({ type, ...payload }, () => void chrome.runtime.lastError); } catch (_) {}
  };
  const clean = value => String(value || "")
    .replace(/\u200b/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const visible = node => !!node && node.getClientRects().length > 0;

  const rootSelectors = [
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
  const bodySelectors = [
    '[id^="message-body-"]',
    '[data-tid="message-content"]',
    '[data-tid="message-body"]',
    '[data-tid="bubble-text"]',
    '[data-tid="chat-pane-message-text"]',
    '[data-tid="message-bubble"]',
    '[data-tid^="message-content-"]',
    '[data-tid^="messageBody"]',
    '.ts-message-content',
    '.message-body',
    '[aria-label^="Message by "]'
  ];
  const itemSelectors = [
    '[data-tid="message-group-container"]',
    '[data-tid="message-list-item"]',
    '[data-tid="message-list-item-container"]',
    'div[role="listitem"]',
    '[aria-label^="Message by "]'
  ];
  const authorSelectors = [
    '[data-tid="message-author-name"]',
    '[data-tid="author-name"]',
    '[id^="message-author-"]',
    '[data-tid="message-group-author"]',
    '[data-tid="message-group-header"]'
  ];
  const timeSelectors = [
    '[id^="timestamp-"]',
    '[data-tid="message-time"]',
    '[data-tid="message-timestamp"]',
    'time[datetime]',
    'time'
  ];

  function countMessageBodies(root) {
    let count = 0;
    for (const selector of bodySelectors) count += root.querySelectorAll(selector).length;
    if (root.matches?.('[aria-label^="Message by "]')) count++;
    return count;
  }

  function findRoot() {
    const candidates = [];
    for (const selector of rootSelectors) {
      for (const node of document.querySelectorAll(selector)) {
        if (!visible(node)) continue;
        const score = countMessageBodies(node);
        if (score > 0) candidates.push({ node, score, selector });
      }
    }
    candidates.sort((a, b) => b.score - a.score);
    return candidates[0] || null;
  }

  function messageBodies(node) {
    if (!node) return [];
    for (const selector of bodySelectors) {
      const found = [];
      if (node.matches?.(selector) && visible(node)) found.push(node);
      found.push(...Array.from(node.querySelectorAll(selector)).filter(visible));
      if (found.length) {
        return found.filter((entry, index) =>
          !found.some((other, otherIndex) => otherIndex !== index && other.contains(entry))
        );
      }
    }
    return [];
  }

  function findItems(root) {
    for (const selector of itemSelectors) {
      const found = Array.from(root.querySelectorAll(selector))
        .filter(node => visible(node) && messageBodies(node).length);
      if (found.length) {
        return found.filter(node => !found.some(other => other !== node && other.contains(node)));
      }
    }
    return messageBodies(root).map(body =>
      body.closest(itemSelectors.join(",")) || body.parentElement
    ).filter((node, index, all) => node && all.indexOf(node) === index);
  }

  function readMeta(start, selectors) {
    let node = start;
    for (let depth = 0; node && depth < 7; depth++, node = node.parentElement) {
      for (const selector of selectors) {
        const found = node.matches?.(selector) ? node : node.querySelector(selector);
        if (!found) continue;
        const value = clean(found.getAttribute("aria-label") || found.getAttribute("title") || found.innerText || found.textContent);
        if (value) return value;
      }
      const label = node.getAttribute?.("aria-label") || "";
      const match = label.match(/^Message by\s+([^,]+)/i);
      if (match) return clean(match[1]);
    }
    return "";
  }

  function currentUserName() {
    for (const node of document.querySelectorAll("[data-tid], [role='treeitem'], [role='listitem']")) {
      if (!visible(node)) continue;
      const match = clean(node.innerText || node.textContent).match(/^(.{2,80}?)\s*\(You\)$/i);
      if (match) return clean(match[1]);
    }
    return "";
  }

  function redact(text) {
    return clean(text)
      .replace(/\b(password|passcode|passwd|pwd|otp|one[- ]time code|api[ _-]?key|access[ _-]?token|secret|2fa)\b\s*(?:is\s*)?[:=]?\s*([^\s,;]+)/gi, "$1 [hidden]")
      .replace(/\bBearer\s+[A-Za-z0-9._~+\/-]+=*/gi, "Bearer [hidden]")
      .replace(/\b(?:sk-[A-Za-z0-9_-]{16,}|gh[pousr]_[A-Za-z0-9]{20,})\b/g, "[access token hidden]")
      .replace(/\b(?=[A-Za-z0-9!@#$%^&*_-]{12,}\b)(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[!@#$%^&*_-])[A-Za-z0-9!@#$%^&*_-]+\b/g, "[possible password hidden]");
  }

  function isOwn(item, body, author, me) {
    if (/^(you|me)$/i.test(author)) return true;
    if (me && author.toLowerCase() === me.toLowerCase()) return true;
    let node = body;
    for (let depth = 0; node && depth < 8; depth++, node = node.parentElement) {
      const value = [
        typeof node.className === "string" ? node.className : "",
        node.getAttribute("data-tid"),
        node.getAttribute("data-message-direction"),
        node.getAttribute("data-is-own-message"),
        node.getAttribute("aria-label")
      ].filter(Boolean).join(" ").toLowerCase();
      if (/outgoing|message-out|own-message|self-message|sent-by-me|from-me|you sent|you said/.test(value)) return true;
      if (node === item) break;
    }
    return false;
  }

  function extract(root) {
    const items = findItems(root);
    const output = [];
    const me = currentUserName();

    items.forEach((item, index) => {
      const bodies = messageBodies(item);
      const groupAuthor = readMeta(item, authorSelectors);

      bodies.forEach(body => {
        const label = body.getAttribute("aria-label") || item.getAttribute("aria-label") || "";
        const labelMatch = label.match(/^Message by\s+([^,]+)/i);
        const author = readMeta(body, authorSelectors) || groupAuthor || (labelMatch && labelMatch[1]) || "";
        const sender = clean(author);
        if (!sender) return;

        const nested = body.querySelector('[id^="content-"]');
        const source = nested || body;
        const clone = source.cloneNode(true);
        clone.querySelectorAll(
          "button,[role='button'],svg,script,style,nav,[aria-hidden='true'],[data-tid*='reaction'],[data-testid*='reaction']"
        ).forEach(node => node.remove());

        const raw = clean(clone.innerText || clone.textContent);
        if (!raw || /^Message by\s+/i.test(raw)) return;

        const stamp = readMeta(body, timeSelectors) || readMeta(item, timeSelectors);
        const id = body.id || body.getAttribute("data-message-id") || item.getAttribute("data-message-id") || "";
        const key = id || [sender, stamp, raw].join("|");

        output.push({
          key,
          sender,
          text: redact(raw),
          timestamp: stamp,
          own: isOwn(item, body, sender, me),
          index
        });
      });
    });
    return { messages: output, itemCount: items.length };
  }

  function chatKey(info) {
    const selectors = [
      '[data-tid="entity-header"]',
      '[data-tid="thread-header"]',
      '[data-tid="chat-header"]',
      '[data-tid="chat-title"]',
      '[data-tid="chat-name"]',
      '[data-tid="conversation-title"]',
      'header [role="heading"]',
      'h1'
    ];
    let title = "";
    for (const selector of selectors) {
      const element = document.querySelector(selector);
      if (element) {
        title = clean(element.innerText || element.textContent);
        if (title) break;
      }
    }
    return location.pathname + "|" + title;
  }

  function stop() {
    state.enabled = false;
    state.running = false;
    state.observer?.disconnect();
    state.waitingObserver?.disconnect();
    if (state.timer) clearTimeout(state.timer);
    if (state.poller) clearInterval(state.poller);
    state.observer = state.waitingObserver = state.poller = null;
    state.root = null;
  }

  function start() {
    if (!/^https:\/\/(teams\.cloud\.microsoft|teams\.microsoft\.com)\//i.test(location.href)) return;
    if (state.running) return;
    state.enabled = true;
    state.running = true;
    state.root = null;
    state.seen.clear();

    function scheduleScan() {
      if (!state.running || state.timer) return;
      state.timer = setTimeout(() => {
        state.timer = null;
        scan();
      }, 180);
    }

    function attach(info) {
      state.root = info.node;
      state.rootSelector = info.selector;
      state.chatKey = chatKey(info);
      const baseline = extract(info.node);
      state.seen = new Set(baseline.messages.map(message => message.key));
      state.observer?.disconnect();
      state.observer = new MutationObserver(scheduleScan);
      state.observer.observe(info.node, { childList: true, subtree: true, characterData: true });
      state.waitingObserver?.disconnect();
      send("TEAMS_LIVE_STATUS", {
        running: true,
        text: "Monitoring Teams chat. Existing messages skipped: " + baseline.messages.length +
          "; message containers found: " + baseline.itemCount + "."
      });
    }

    function scan() {
      if (!state.running) return;
      const info = findRoot();
      if (!info) return;

      if (info.node !== state.root || info.selector !== state.rootSelector || chatKey(info) !== state.chatKey) {
        attach(info);
        return;
      }

      const batch = extract(info.node);
      if (Date.now() < state.suppressUntil) {
        state.seen = new Set(batch.messages.map(message => message.key));
        return;
      }
      const unseen = batch.messages.filter(message => !state.seen.has(message.key));
      // A large replacement usually means a chat was switched or history was loaded.
      // Re-baseline rather than speaking a whole previous conversation as new.
      if (unseen.length > 5 && unseen.length > batch.messages.length * 0.4) {
        state.seen = new Set(batch.messages.map(message => message.key));
        send("TEAMS_LIVE_STATUS", {
          running: true,
          text: "Chat content changed. Existing messages were skipped; listening for the next new message."
        });
        return;
      }

      unseen.forEach(message => {
        state.seen.add(message.key);
        if (message.own) return;
        send("TEAMS_LIVE_MESSAGE", {
          sender: message.sender,
          text: message.text,
          timestamp: message.timestamp,
          messageId: message.key
        });
      });

      if (state.seen.size > 3000) state.seen = new Set(Array.from(state.seen).slice(-1800));
    }

    state.waitingObserver = new MutationObserver(scheduleScan);
    state.waitingObserver.observe(document.documentElement, { childList: true, subtree: true });
    scan();
    state.poller = setInterval(scan, 1200);
    send("TEAMS_LIVE_STATUS", {
      running: true,
      text: "Live reader is ON in this Teams tab. Waiting for the chat message list if it is not loaded yet."
    });
  }

  document.addEventListener("click", event => {
    const target = event.target instanceof Element ? event.target : event.target?.parentElement;
    if (!target) return;
    if (target.closest('[role="treeitem"], [data-tid="chat-list-item"], [data-tid="chat-pane-list"]')) {
      state.suppressUntil = Date.now() + 2800;
      if (state.timer) clearTimeout(state.timer);
      state.timer = setTimeout(() => {
        state.timer = null;
        if (state.running) {
          const info = findRoot();
          if (info) {
            const baseline = extract(info.node);
            state.root = info.node;
            state.rootSelector = info.selector;
            state.chatKey = chatKey(info);
            state.seen = new Set(baseline.messages.map(message => message.key));
            send("TEAMS_LIVE_STATUS", {
              running: true,
              text: "Chat switched. Previous messages skipped: " + baseline.messages.length + "."
            });
          }
        }
      }, 2800);
    }
  }, true);

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local" || !changes.teamsLiveEnabled) return;
    if (changes.teamsLiveEnabled.newValue) start();
    else stop();
  });

  chrome.runtime.onMessage.addListener(message => {
    if (message?.type === "TEAMS_SET_ENABLED") {
      if (message.enabled) start();
      else stop();
    }
  });

  chrome.storage.local.get("teamsLiveEnabled").then(saved => {
    if (saved.teamsLiveEnabled) start();
  }).catch(() => {});
})();
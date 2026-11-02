const $ = id => document.getElementById(id);
let pairs = [];
let selectedId = null;
let voices = [];
let speechRun = 0;

function setStatus(message) {
  $("status").textContent = message;
}

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
          let content = element;
          if (role === "user") {
            content = element.querySelector(".whitespace-pre-wrap") ||
              element.querySelector("[data-message-content]") ||
              element;
          } else {
            content = element.querySelector(".markdown") ||
              element.querySelector("[data-message-content]") ||
              element;
          }

          const clone = content.cloneNode(true);
          clone.querySelectorAll(
            "button, [role='button'], svg, script, style, nav, [aria-hidden='true'], [data-testid*='action'], [data-testid*='toolbar']"
          ).forEach(node => node.remove());

          return clean(clone.innerText || clone.textContent);
        };

        const rawRoleNodes = Array.from(document.querySelectorAll(
          'main [data-message-author-role="user"], main [data-message-author-role="assistant"], [data-message-author-role="user"], [data-message-author-role="assistant"]'
        )).filter(visible);

        let messages = [];
        const normalize = value => clean(value).replace(/\s+/g, " ").toLowerCase();

        if (rawRoleNodes.length) {
          // Drop same-role descendants; they are nested render copies, not messages.
          const roleNodes = rawRoleNodes.filter(node => {
            const role = node.getAttribute("data-message-author-role");
            return !node.parentElement?.closest('[data-message-author-role="' + role + '"]');
          });

          const seenInTurn = new Set();
          roleNodes.forEach((node, index) => {
            const role = node.getAttribute("data-message-author-role");
            const turn = node.closest('[data-testid^="conversation-turn-"]') ||
              node.closest("article") || node.parentElement;
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
loadConversation();

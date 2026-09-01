// Mesmo shim de background.js: Firefox só expõe `browser.*` em páginas de
// extensão.
if (!globalThis.chrome && globalThis.browser) {
  globalThis.chrome = globalThis.browser;
}

const transcriptEl = document.getElementById("transcript");
const summaryEl = document.getElementById("summary");
const newMeetingButton = document.getElementById("newMeeting");
const clearButton = document.getElementById("clear");

function formatTime(value) {
  return new Date(value).toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function reportStorageError(operation, error) {
  console.error(`❌ Falha ${operation}`, error);
}

async function sendStorageCommand(message) {
  const response = await chrome.runtime.sendMessage(message);
  if (!response?.ok) {
    throw new Error(response?.error || "Falha desconhecida no storage");
  }
  return response;
}

async function render() {
  const result = await chrome.storage.local.get({ meetingTranscript: [] });
  const transcript = result.meetingTranscript;

  summaryEl.textContent =
    `${transcript.length} fala${transcript.length === 1 ? "" : "s"} nesta reunião`;

  if (!transcript.length) {
    transcriptEl.innerHTML =
      '<div id="empty">Aguardando legendas do Teams...</div>';
    return;
  }

  transcriptEl.innerHTML = transcript.map(item => `
    <div class="item">
      <div class="meta">
        ${escapeHtml(formatTime(item.time))}
        ·
        <span class="speaker">${escapeHtml(item.speaker)}</span>
      </div>
      <div class="text">${escapeHtml(item.text)}</div>
    </div>
  `).join("");

  transcriptEl.scrollTop = transcriptEl.scrollHeight;
}

newMeetingButton.addEventListener("click", async () => {
  try {
    await sendStorageCommand({
      type: "meetingCopilot:newMeeting",
      nextMeetingId: crypto.randomUUID(),
      endedAt: new Date().toISOString()
    });
    await render();
  } catch (error) {
    reportStorageError("ao iniciar nova reunião", error);
  }
});

clearButton.addEventListener("click", async () => {
  if (!confirm("Limpar a reunião atual?")) return;

  try {
    await sendStorageCommand({
      type: "meetingCopilot:clearTranscript",
      nextMeetingId: crypto.randomUUID()
    });
    await render();
  } catch (error) {
    reportStorageError("ao limpar reunião", error);
  }
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes.meetingTranscript) {
    render().catch((error) => {
      reportStorageError("ao atualizar Side Panel", error);
    });
  }
  if (area === "local" && (changes.meetingState || changes.meetingInsights)) {
    refreshState().catch(() => {});
  }
});

render().catch((error) => {
  reportStorageError("ao carregar Side Panel", error);
});

// ---------------------------------------------------------------------------
// Abas
// ---------------------------------------------------------------------------

for (const tab of document.querySelectorAll("nav button")) {
  tab.addEventListener("click", () => {
    for (const other of document.querySelectorAll("nav button")) {
      other.setAttribute("aria-selected", String(other === tab));
    }
    for (const panel of document.querySelectorAll(".panel")) {
      panel.classList.toggle("active", panel.id === `panel-${tab.dataset.panel}`);
    }
  });
}

// ---------------------------------------------------------------------------
// Estado da reunião + insights
// ---------------------------------------------------------------------------

const insightsEl = document.getElementById("insights");
const stateViewEl = document.getElementById("stateView");

function renderInsights(insights) {
  if (!insights.length) {
    insightsEl.innerHTML = "";
    return;
  }

  insightsEl.innerHTML = insights.map((insight) => `
    <div class="insight" data-id="${escapeHtml(insight.id)}">
      <button data-dismiss="${escapeHtml(insight.id)}" title="Dispensar">×</button>
      <div class="kind">${escapeHtml(insight.kind)}</div>
      <div class="text">${escapeHtml(insight.text)}</div>
      ${insight.why ? `<div class="why">${escapeHtml(insight.why)}</div>` : ""}
    </div>
  `).join("");

  for (const button of insightsEl.querySelectorAll("[data-dismiss]")) {
    button.addEventListener("click", () => {
      chrome.runtime.sendMessage({
        type: "meetingCopilot:dismissInsight",
        id: button.dataset.dismiss
      });
    });
  }
}

function renderStateSection(title, items, formatter) {
  if (!items?.length) return "";
  return `<h2>${title}</h2><ul>${items.map(formatter).join("")}</ul>`;
}

function renderState(state) {
  if (!state || !state.summary) {
    stateViewEl.innerHTML = '<div class="empty">Aguardando falas suficientes para montar o estado.</div>';
    return;
  }

  const item = (text) => `<li>${escapeHtml(text)}</li>`;
  const commitment = (c) => `<li>${escapeHtml(c.what)}${c.who ? ` — <em>${escapeHtml(c.who)}</em>` : ""}</li>`;

  stateViewEl.innerHTML = `
    <h2>Resumo</h2>
    <p>${escapeHtml(state.summary)}</p>
    ${renderStateSection("Tópicos", state.topics, (t) => item(t.text))}
    ${renderStateSection("Decisões", state.decisions, (t) => item(t.text))}
    ${renderStateSection("Riscos", state.risks, (t) => item(t.text))}
    ${renderStateSection("Perguntas abertas", state.openQuestions, (t) => item(t.text))}
    ${renderStateSection("Compromissos", state.commitments, commitment)}
  `;
}

async function refreshState() {
  const result = await chrome.storage.local.get({
    meetingState: null,
    meetingInsights: []
  });
  renderInsights(result.meetingInsights || []);
  renderState(result.meetingState);
}

refreshState().catch(() => {});

// ---------------------------------------------------------------------------
// Chat com streaming
// ---------------------------------------------------------------------------

const shortcutsEl = document.getElementById("shortcuts");
const chatLogEl = document.getElementById("chatLog");
const chatForm = document.getElementById("chatForm");
const chatInput = document.getElementById("chatInput");

let chatPort = null;
let currentAssistantEl = null;

function ensureChatPort() {
  if (chatPort) return chatPort;

  chatPort = chrome.runtime.connect({ name: "meetingCopilot:chat" });
  chatPort.onMessage.addListener((message) => {
    if (message.type === "token") {
      if (currentAssistantEl) {
        currentAssistantEl.textContent += message.text;
        chatLogEl.scrollTop = chatLogEl.scrollHeight;
      }
      return;
    }

    if (message.type === "done") {
      if (currentAssistantEl && !currentAssistantEl.textContent) {
        currentAssistantEl.textContent = message.text || "(sem resposta)";
      }
      currentAssistantEl = null;
      return;
    }

    if (message.type === "error") {
      if (currentAssistantEl) {
        currentAssistantEl.textContent = `⚠️ ${message.error}`;
        if (message.hint) currentAssistantEl.textContent += `\n${message.hint}`;
      }
      currentAssistantEl = null;
    }
  });
  chatPort.onDisconnect.addListener(() => { chatPort = null; });

  return chatPort;
}

function addTurn(role, text) {
  const el = document.createElement("div");
  el.className = `turn ${role}`;
  el.textContent = text;
  chatLogEl.appendChild(el);
  chatLogEl.scrollTop = chatLogEl.scrollHeight;
  return el;
}

function askQuestion(prompt) {
  addTurn("you", prompt);
  currentAssistantEl = addTurn("assistant", "");
  ensureChatPort().postMessage({ type: "ask", prompt });
}

chatForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const value = chatInput.value.trim();
  if (!value) return;
  chatInput.value = "";
  askQuestion(value);
});

function loadShortcuts() {
  // Lista replicada de meeting-state.js — o side panel não importa módulos.
  const shortcuts = [
    { label: "Perguntar", prompt: "Que pergunta forte eu poderia fazer agora nesta reunião?" },
    { label: "Responder", prompt: "Como eu poderia responder ao último ponto levantado?" },
    { label: "Explicar", prompt: "Explique o último ponto técnico discutido, de forma simples." },
    { label: "Contestar", prompt: "Qual é o contra-argumento mais forte ao que foi dito?" },
    { label: "Riscos", prompt: "Quais riscos ou pontos cegos aparecem no que foi discutido?" },
    { label: "Resumo", prompt: "Resuma a reunião até agora em tópicos objetivos." },
    { label: "Minha vez", prompt: "É a minha vez de falar. O que vale eu dizer agora, em 2 ou 3 frases?" }
  ];

  shortcutsEl.innerHTML = shortcuts.map((s, index) =>
    `<button data-shortcut="${index}">${escapeHtml(s.label)}</button>`
  ).join("");

  for (const button of shortcutsEl.querySelectorAll("[data-shortcut]")) {
    button.addEventListener("click", () => {
      askQuestion(shortcuts[Number(button.dataset.shortcut)].prompt);
    });
  }
}

loadShortcuts();

// ---------------------------------------------------------------------------
// Configuração
// ---------------------------------------------------------------------------

const providerEl = document.getElementById("provider");
const ollamaUrlEl = document.getElementById("ollamaUrl");
const modelEl = document.getElementById("model");
const modelNoteEl = document.getElementById("modelNote");
const externalKeyEl = document.getElementById("externalKey");
const externalModelEl = document.getElementById("externalModel");
const meetingStateEnabledEl = document.getElementById("meetingStateEnabled");
const proactiveEl = document.getElementById("proactive");
const blockSizeEl = document.getElementById("blockSize");
const audioFallbackEl = document.getElementById("audioFallback");
const whisperUrlEl = document.getElementById("whisperUrl");
const retentionDaysEl = document.getElementById("retentionDays");
const debugEl = document.getElementById("debug");
const configStatusEl = document.getElementById("configStatus");
const refreshModelsButton = document.getElementById("refreshModels");
const forgetButton = document.getElementById("forget");

function applySettingsToForm(settings) {
  providerEl.value = settings.provider;
  ollamaUrlEl.value = settings.ollamaUrl;
  externalModelEl.value = settings.externalModel || "";
  meetingStateEnabledEl.checked = settings.meetingStateEnabled;
  proactiveEl.checked = settings.proactive;
  blockSizeEl.value = settings.proactiveBlockSize;
  audioFallbackEl.checked = settings.audioFallback;
  whisperUrlEl.value = settings.whisperUrl;
  retentionDaysEl.value = settings.retentionDays;
  debugEl.checked = settings.debug;

  if (settings.model) {
    modelEl.innerHTML = `<option value="${escapeHtml(settings.model)}">${escapeHtml(settings.model)}</option>`;
  }
}

async function loadConfig() {
  const response = await chrome.runtime.sendMessage({ type: "meetingCopilot:getSettings" });
  if (response?.ok) applySettingsToForm(response.settings);
}

async function persistSetting(patch) {
  const response = await chrome.runtime.sendMessage({
    type: "meetingCopilot:setSettings",
    settings: patch
  });
  if (!response?.ok) {
    configStatusEl.textContent = `Falha ao salvar: ${response?.error || "erro desconhecido"}`;
    return;
  }
  configStatusEl.textContent = "";
}

async function detectModels() {
  modelNoteEl.textContent = "Procurando modelos no Ollama...";
  const response = await chrome.runtime.sendMessage({ type: "meetingCopilot:listModels" });

  if (!response?.ok) {
    modelNoteEl.textContent = response?.error || "Não foi possível listar modelos.";
    if (response?.hint) modelNoteEl.textContent += ` ${response.hint}`;
    return;
  }

  if (!response.models.length) {
    modelNoteEl.textContent = "Nenhum modelo instalado. Rode: ollama pull <modelo>";
    return;
  }

  modelEl.innerHTML = response.models
    .map((m) => `<option value="${escapeHtml(m.name)}">${escapeHtml(m.name)}</option>`)
    .join("");
  modelNoteEl.textContent = `${response.models.length} modelo(s) encontrado(s).`;
  await persistSetting({ model: modelEl.value });
}

providerEl.addEventListener("change", () => persistSetting({ provider: providerEl.value }));
ollamaUrlEl.addEventListener("change", () => persistSetting({ ollamaUrl: ollamaUrlEl.value }));
modelEl.addEventListener("change", () => persistSetting({ model: modelEl.value }));
const EXTERNAL_PROVIDER_ORIGINS = {
  anthropic: "https://api.anthropic.com/*",
  openai: "https://api.openai.com/*"
};

externalKeyEl.addEventListener("change", async () => {
  const key = externalKeyEl.value;
  externalKeyEl.value = "";
  if (!key) return;

  // O host do provider externo é optional_host_permissions: só é concedido
  // quando o usuário de fato entra com uma chave — nunca na instalação.
  const origin = EXTERNAL_PROVIDER_ORIGINS[providerEl.value];
  if (origin && chrome.permissions?.request) {
    const granted = await chrome.permissions.request({ origins: [origin] });
    if (!granted) {
      configStatusEl.textContent = "Permissão negada — provider externo continua desligado.";
      return;
    }
  }

  await persistSetting({ externalApiKey: key, externalProvider: providerEl.value });
});
externalModelEl.addEventListener("change", () => persistSetting({ externalModel: externalModelEl.value }));
meetingStateEnabledEl.addEventListener("change", () => persistSetting({ meetingStateEnabled: meetingStateEnabledEl.checked }));
proactiveEl.addEventListener("change", () => persistSetting({ proactive: proactiveEl.checked }));
blockSizeEl.addEventListener("change", () => persistSetting({ proactiveBlockSize: Number(blockSizeEl.value) }));
audioFallbackEl.addEventListener("change", () => persistSetting({ audioFallback: audioFallbackEl.checked }));
whisperUrlEl.addEventListener("change", () => persistSetting({ whisperUrl: whisperUrlEl.value }));
retentionDaysEl.addEventListener("change", () => persistSetting({ retentionDays: Number(retentionDaysEl.value) }));
debugEl.addEventListener("change", () => persistSetting({ debug: debugEl.checked }));
refreshModelsButton.addEventListener("click", () => detectModels().catch((e) => {
  modelNoteEl.textContent = String(e?.message ?? e);
}));

forgetButton.addEventListener("click", async () => {
  if (!confirm("Apagar transcript, histórico, estado e insights deste navegador?")) return;
  await chrome.runtime.sendMessage({ type: "meetingCopilot:forgetEverything" });
  configStatusEl.textContent = "Dados apagados.";
  await render();
  await refreshState();
});

loadConfig().catch(() => {});

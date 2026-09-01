// Firefox só expõe `browser.*` (promise-based) em páginas de extensão;
// aliasar para `chrome` deixa o resto do arquivo idêntico nos dois
// navegadores, sem espalhar `browser ?? chrome` em cada chamada.
if (!globalThis.chrome && globalThis.browser) {
  globalThis.chrome = globalThis.browser;
}

// `importScripts` não existe no sandbox de teste (vm.createContext); em
// produção o service worker sempre tem a função global.
if (typeof importScripts === "function") {
  try {
    importScripts("llm.js", "meeting-state.js");
  } catch (error) {
    console.error("❌ Falha ao carregar camadas de LLM", error);
  }
}

// chrome.sidePanel não existe no Firefox (usa sidebar_action, declarado no
// manifest e aberto automaticamente — nada a fazer aqui nesse caso).
chrome.sidePanel?.setPanelBehavior?.({ openPanelOnActionClick: true })
  ?.catch((error) => console.error(error));

const LOG_PREFIX = "[MeetingCopilot]";
const MAX_STORAGE_ATTEMPTS = 3;
const TEAMS_URL_PATTERN = "https://teams.microsoft.com/*";
const TEAMS_URL_PREFIX = "https://teams.microsoft.com/";
const CONTENT_SCRIPT_FILES = ["transcript.js", "content.js"];
let storageWriteQueue = Promise.resolve();

// Executada dentro da aba, no mesmo isolated world do content script.
function meetingCopilotProbe() {
  return globalThis.__MEETING_COPILOT_PROBE__?.() === true;
}

async function hasLiveCapture(tabId) {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: meetingCopilotProbe
    });
    return results?.some((entry) => entry?.result === true) === true;
  } catch (error) {
    // Aba sem permissão, descarregada ou em processo de navegação:
    // trata como sem captura e deixa a injeção decidir.
    return false;
  }
}

// O Chrome NÃO substitui content scripts já injetados quando a extensão é
// recarregada/atualizada — o script antigo fica órfão com o contexto
// invalidado. Reinjetar pelo service worker é a única recuperação automática.
let sweepInFlight = null;
let sweepQueued = null;

// Vários gatilhos (boot do worker, onInstalled, onUpdated) disparam quase
// juntos num reload. Sem isto, a mesma aba levaria injeções concorrentes e
// reiniciaria a captura mais de uma vez.
function reinjectTeamsTabs(reason) {
  if (!sweepInFlight) return startSweep(reason);

  if (!sweepQueued) {
    sweepQueued = sweepInFlight.then(() => {
      sweepQueued = null;
      return startSweep(reason);
    });
  }

  return sweepQueued;
}

function startSweep(reason) {
  sweepInFlight = runSweep(reason).finally(() => {
    sweepInFlight = null;
  });
  return sweepInFlight;
}

async function runSweep(reason) {
  if (!chrome.tabs?.query || !chrome.scripting?.executeScript) {
    return { injected: [], skipped: [] };
  }

  let tabs = [];
  try {
    tabs = await chrome.tabs.query({ url: TEAMS_URL_PATTERN });
  } catch (error) {
    console.error(`❌ Falha ao listar abas do Teams (${reason})`, error);
    return { injected: [], skipped: [] };
  }

  const injected = [];
  const skipped = [];

  for (const tab of tabs) {
    if (typeof tab?.id !== "number" || tab.id < 0) continue;

    if (await hasLiveCapture(tab.id)) {
      skipped.push(tab.id);
      continue;
    }

    try {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: CONTENT_SCRIPT_FILES
      });
      injected.push(tab.id);
    } catch (error) {
      console.error(`❌ Falha ao reinjetar na aba ${tab.id} (${reason})`, error);
    }
  }

  // O service worker acorda a cada fala salva e reexecuta este sweep; logar
  // só quando algo foi de fato reinjetado mantém o console legível.
  if (injected.length) {
    console.log?.(`${LOG_PREFIX} reinjeção`, { reason, injected, skipped });
  }

  return { injected, skipped };
}

chrome.runtime.onInstalled?.addListener?.(({ reason } = {}) => {
  reinjectTeamsTabs(`onInstalled:${reason}`);
});

chrome.runtime.onStartup?.addListener?.(() => {
  reinjectTeamsTabs("onStartup");
  applyRetention?.().catch(() => {});
});

chrome.tabs?.onUpdated?.addListener?.((_tabId, changeInfo, tab) => {
  if (changeInfo?.status !== "complete") return;
  if (!String(tab?.url || "").startsWith(TEAMS_URL_PREFIX)) return;
  reinjectTeamsTabs("aba do Teams carregada");
});

// Também no boot do service worker: é o gatilho que sobrevive ao "Recarregar"
// de extensão sem compactação, independente de onInstalled disparar.
reinjectTeamsTabs("service worker iniciado");

async function withStorageRetry(operation, label) {
  let lastError;

  for (let attempt = 1; attempt <= MAX_STORAGE_ATTEMPTS; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      console.error(
        `❌ Falha ${label} (tentativa ${attempt}/${MAX_STORAGE_ATTEMPTS})`,
        error
      );
    }
  }

  throw lastError;
}

function enqueueStorageOperation(label, operation) {
  const result = storageWriteQueue.then(() =>
    withStorageRetry(operation, label)
  );
  storageWriteQueue = result.catch(() => {});

  return result
    .then((value) => ({ ok: true, ...(value || {}) }))
    .catch((error) => ({
      ok: false,
      error: String(error?.message ?? error)
    }));
}

async function appendTranscript(item) {
  const result = await chrome.storage.local.get({
    meetingTranscript: [],
    meetingHistory: [],
    meetingClosedSessions: [],
    meetingSessionId: null
  });
  const transcript = result.meetingTranscript || [];
  const history = result.meetingHistory || [];
  const closedSessions = result.meetingClosedSessions || [];

  const alreadySaved = transcript.some((saved) => saved.id === item.id) ||
    history.some(({ transcript: archived = [] }) =>
      archived.some((saved) => saved.id === item.id)
    );
  if (alreadySaved) return;

  // Sessão não resolvida no content script: preserva a fala na reunião atual
  // em vez de descartá-la em silêncio.
  if (!item.meetingId) {
    console.warn?.(`${LOG_PREFIX} fala sem meetingId anexada à sessão atual`, {
      id: item.id
    });
    await chrome.storage.local.set({
      meetingTranscript: [...transcript, item]
    });
    return;
  }

  if (item.meetingId === result.meetingSessionId) {
    await chrome.storage.local.set({
      meetingTranscript: [...transcript, item]
    });
    return;
  }

  let archivedMeeting = history.find(({ id, sessionId }) =>
    (sessionId || id) === item.meetingId
  );
  if (!archivedMeeting) {
    const closedSession = closedSessions.find(({ id }) =>
      id === item.meetingId
    );
    if (!closedSession) return;

    archivedMeeting = {
      id: closedSession.id,
      endedAt: closedSession.endedAt,
      transcript: []
    };
    history.push(archivedMeeting);
  }

  archivedMeeting.transcript = [...(archivedMeeting.transcript || []), item];
  await chrome.storage.local.set({ meetingHistory: history });
}

async function startNewMeeting({ nextMeetingId, endedAt }) {
  const result = await chrome.storage.local.get({
    meetingTranscript: [],
    meetingHistory: [],
    meetingClosedSessions: [],
    meetingSessionId: null
  });
  const transcript = result.meetingTranscript || [];
  const history = result.meetingHistory || [];
  const closedSessions = result.meetingClosedSessions || [];
  const previousMeetingId = result.meetingSessionId || `legacy-${endedAt}`;

  if (previousMeetingId === nextMeetingId) return;

  if (
    previousMeetingId &&
    !closedSessions.some(({ id }) => id === previousMeetingId)
  ) {
    closedSessions.push({ id: previousMeetingId, endedAt });
  }

  if (
    transcript.length &&
    previousMeetingId &&
    !history.some(({ id, sessionId }) =>
      (sessionId || id) === previousMeetingId
    )
  ) {
    history.push({
      id: previousMeetingId,
      endedAt,
      transcript
    });
  }

  await chrome.storage.local.set({
    meetingHistory: history,
    meetingClosedSessions: closedSessions,
    meetingTranscript: [],
    meetingSessionId: nextMeetingId
  });
}

async function ensureMeetingSession(candidateMeetingId) {
  const result = await chrome.storage.local.get({ meetingSessionId: null });
  const meetingId = result.meetingSessionId || candidateMeetingId;

  if (!result.meetingSessionId) {
    await chrome.storage.local.set({ meetingSessionId: meetingId });
  }

  return { meetingId };
}

function handleStorageMessage(message) {
  if (message?.type === "meetingCopilot:appendTranscript") {
    const saved = enqueueStorageOperation(
      "ao salvar transcript",
      () => appendTranscript(message.item)
    );
    // Análise roda em bloco e fora do caminho crítico da persistência.
    saved.then(() => runAnalysis?.("nova fala")).catch(() => {});
    return saved;
  }

  if (message?.type === "meetingCopilot:newMeeting") {
    return enqueueStorageOperation(
      "ao iniciar nova reunião",
      () => startNewMeeting(message)
    );
  }

  if (message?.type === "meetingCopilot:clearTranscript") {
    return enqueueStorageOperation(
      "ao limpar transcript",
      () => chrome.storage.local.set({
        meetingTranscript: [],
        meetingSessionId: message.nextMeetingId
      })
    );
  }

  if (message?.type === "meetingCopilot:getMeetingSession") {
    return enqueueStorageOperation(
      "ao carregar sessão da reunião",
      () => ensureMeetingSession(message.candidateMeetingId)
    );
  }

  return undefined;
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  const response = handleStorageMessage(message) ??
    handleAiMessage?.(message) ??
    handleAudioMessage?.(message, sender);
  if (!response) return undefined;

  response.then(sendResponse);
  return true;
});

// ---------------------------------------------------------------------------
// Camada de IA local: settings, Meeting State e copiloto proativo.
// Tudo isolado da captura — se o Ollama cair, as legendas continuam salvando.
// ---------------------------------------------------------------------------

const SETTINGS_KEY = "meetingCopilotSettings";
const STATE_KEY = "meetingState";
const INSIGHTS_KEY = "meetingInsights";
const MAX_INSIGHTS = 20;

function llmApi() {
  return globalThis.MeetingLlm || null;
}

function stateApi() {
  return globalThis.MeetingState || null;
}

async function loadSettings() {
  const api = llmApi();
  const stored = await chrome.storage.local.get({ [SETTINGS_KEY]: {} });
  const raw = stored[SETTINGS_KEY] || {};
  return api ? api.normalizeSettings(raw) : raw;
}

async function saveSettings(patch) {
  const current = await loadSettings();
  const api = llmApi();
  const merged = { ...current, ...(patch || {}) };
  const next = api ? api.normalizeSettings(merged) : merged;
  await chrome.storage.local.set({ [SETTINGS_KEY]: next });
  return next;
}

async function createClient() {
  const api = llmApi();
  if (!api) throw new Error("Camada de LLM indisponível.");
  return api.createClient({ settings: await loadSettings() });
}

// --- Meeting State + insights ---------------------------------------------

let analysisInFlight = false;

async function readTranscript() {
  const result = await chrome.storage.local.get({ meetingTranscript: [] });
  return result.meetingTranscript || [];
}

async function runAnalysis(reason) {
  const states = stateApi();
  if (!states || analysisInFlight) return;

  const settings = await loadSettings();
  if (!settings.meetingStateEnabled && !settings.proactive) return;

  const transcript = await readTranscript();
  const stored = await chrome.storage.local.get({ [STATE_KEY]: null });
  const previous = stored[STATE_KEY] || states.emptyState();

  if (!states.shouldAnalyze({
    transcriptLength: transcript.length,
    analyzedCount: previous.segmentCount || 0,
    blockSize: settings.proactiveBlockSize
  })) {
    return;
  }

  analysisInFlight = true;

  try {
    const client = await createClient();

    if (settings.meetingStateEnabled) {
      const prompt = states.buildStatePrompt(transcript, previous);
      const raw = await client.chat({ ...prompt });
      const next = states.parseState(raw, {
        transcript,
        segmentCount: transcript.length,
        now: new Date().toISOString()
      });

      if (next) {
        await chrome.storage.local.set({ [STATE_KEY]: next });
        console.log(`${LOG_PREFIX} meeting state atualizado`, {
          reason,
          falas: transcript.length
        });
      }
    }

    if (settings.proactive) {
      await runInsights(client, transcript, settings);
    }
  } catch (error) {
    // Falha de modelo nunca derruba a captura: apenas registra.
    console.warn(`${LOG_PREFIX} análise indisponível`, String(error?.message ?? error));
  } finally {
    analysisInFlight = false;
  }
}

async function runInsights(client, transcript, settings) {
  const states = stateApi();
  const block = Math.max(settings.proactiveBlockSize, 3);
  const recent = transcript.slice(-block);
  if (!recent.length) return;

  const stored = await chrome.storage.local.get({
    [INSIGHTS_KEY]: [],
    [STATE_KEY]: null
  });
  const existing = stored[INSIGHTS_KEY] || [];
  const prompt = states.buildInsightPrompt(recent, stored[STATE_KEY]);
  const raw = await client.chat({ ...prompt });
  const fresh = states.parseInsights(raw, {
    transcript,
    seen: existing.map(({ text }) => text)
  });

  if (!fresh.length) return;

  const stamped = fresh.map((insight) => ({
    ...insight,
    id: `insight-${Date.now()}-${Math.round(Math.random() * 1e6)}`,
    createdAt: new Date().toISOString()
  }));

  await chrome.storage.local.set({
    [INSIGHTS_KEY]: [...existing, ...stamped].slice(-MAX_INSIGHTS)
  });
  console.log(`${LOG_PREFIX} insights`, { novos: stamped.length });
}

// --- Retenção ---------------------------------------------------------------

async function applyRetention() {
  const settings = await loadSettings();
  if (!settings.retentionDays) return;

  const cutoff = Date.now() - settings.retentionDays * 86400000;
  const stored = await chrome.storage.local.get({ meetingHistory: [] });
  const history = stored.meetingHistory || [];
  const kept = history.filter(({ endedAt }) => {
    const time = Date.parse(endedAt);
    return Number.isNaN(time) || time >= cutoff;
  });

  if (kept.length !== history.length) {
    await chrome.storage.local.set({ meetingHistory: kept });
    console.log(`${LOG_PREFIX} retenção aplicada`, {
      removidas: history.length - kept.length
    });
  }
}

async function forgetEverything() {
  await chrome.storage.local.remove([
    "meetingTranscript",
    "meetingHistory",
    "meetingClosedSessions",
    "meetingSessionId",
    STATE_KEY,
    INSIGHTS_KEY
  ]);
}

// --- Mensagens de IA --------------------------------------------------------

function handleAiMessage(message) {
  if (message?.type === "meetingCopilot:getSettings") {
    return loadSettings().then((settings) => ({ ok: true, settings }));
  }

  if (message?.type === "meetingCopilot:setSettings") {
    return saveSettings(message.settings)
      .then((settings) => ({ ok: true, settings }))
      .catch((error) => ({ ok: false, error: String(error?.message ?? error) }));
  }

  if (message?.type === "meetingCopilot:listModels") {
    return createClient()
      .then((client) => client.listModels())
      .then((models) => ({ ok: true, models }))
      .catch((error) => ({
        ok: false,
        error: String(error?.message ?? error),
        hint: error?.hint || ""
      }));
  }

  if (message?.type === "meetingCopilot:dismissInsight") {
    return chrome.storage.local.get({ [INSIGHTS_KEY]: [] })
      .then(({ [INSIGHTS_KEY]: insights }) => chrome.storage.local.set({
        [INSIGHTS_KEY]: (insights || []).filter(({ id }) => id !== message.id)
      }))
      .then(() => ({ ok: true }));
  }

  if (message?.type === "meetingCopilot:forgetEverything") {
    return forgetEverything().then(() => ({ ok: true }));
  }

  if (message?.type === "meetingCopilot:analyzeNow") {
    return runAnalysis("manual").then(() => ({ ok: true }));
  }

  return undefined;
}

// Streaming de chat via Port: tokens chegam no side panel conforme saem.
chrome.runtime.onConnect?.addListener?.((port) => {
  if (port.name !== "meetingCopilot:chat") return;

  let controller = null;

  port.onDisconnect.addListener(() => controller?.abort());

  port.onMessage.addListener(async (request) => {
    if (request?.type === "cancel") {
      controller?.abort();
      return;
    }

    if (request?.type !== "ask") return;

    controller = new AbortController();

    try {
      const states = stateApi();
      const transcript = await readTranscript();
      const stored = await chrome.storage.local.get({ [STATE_KEY]: null });
      const prompt = states.buildChatPrompt(
        request.prompt,
        transcript,
        stored[STATE_KEY]
      );
      const client = await createClient();

      const full = await client.chat({
        ...prompt,
        signal: controller.signal,
        onToken: (token) => {
          try {
            port.postMessage({ type: "token", text: token });
          } catch (error) {
            controller?.abort();
          }
        }
      });

      port.postMessage({ type: "done", text: full });
    } catch (error) {
      port.postMessage({
        type: "error",
        error: String(error?.message ?? error),
        hint: error?.hint || ""
      });
    } finally {
      controller = null;
    }
  });
});

// ---------------------------------------------------------------------------
// Fallback de áudio (Whisper): só entra quando as legendas falham e o
// usuário ligou a opção. Nunca compete com o áudio da reunião — o
// documento offscreen escuta a aba E devolve o som para a saída.
// ---------------------------------------------------------------------------

const OFFSCREEN_URL = "offscreen.html";
let audioFallbackTabId = null;
let audioFallbackStarting = null;

function hasAudioFallbackApis() {
  // chrome.offscreen/chrome.tabCapture não existem em todo navegador
  // (ex.: Firefox) — sem eles o fallback simplesmente não está disponível.
  return Boolean(chrome.offscreen?.createDocument && chrome.tabCapture?.getMediaStreamId);
}

async function ensureOffscreenDocument() {
  const hasDoc = await chrome.offscreen.hasDocument?.();
  if (hasDoc) return;

  await chrome.offscreen.createDocument({
    url: OFFSCREEN_URL,
    reasons: ["USER_MEDIA"],
    justification: "Capturar áudio da aba do Teams para transcrever via Whisper quando as legendas falham."
  });
}

async function startAudioFallback(tabId) {
  if (!hasAudioFallbackApis()) {
    console.log(`${LOG_PREFIX} fallback de áudio indisponível neste navegador`);
    return false;
  }

  const settings = await loadSettings();
  if (!settings.audioFallback) return false;
  if (audioFallbackTabId === tabId) return true; // já rodando nesta aba

  audioFallbackStarting = (async () => {
    await ensureOffscreenDocument();

    const streamId = await chrome.tabCapture.getMediaStreamId({ targetTabId: tabId });
    const sessionState = await chrome.storage.local.get({ meetingSessionId: null });

    const response = await chrome.runtime.sendMessage({
      type: "meetingCopilot:offscreen:start",
      streamId,
      whisperUrl: settings.whisperUrl,
      meetingId: sessionState.meetingSessionId
    });

    if (!response?.ok) throw new Error(response?.error || "offscreen recusou iniciar");

    audioFallbackTabId = tabId;
    console.log(`${LOG_PREFIX} fallback de áudio iniciado`, { tabId });
  })().catch((error) => {
    console.warn(`${LOG_PREFIX} falha ao iniciar fallback de áudio`, String(error?.message ?? error));
    audioFallbackTabId = null;
  }).finally(() => {
    audioFallbackStarting = null;
  });

  await audioFallbackStarting;
  return audioFallbackTabId === tabId;
}

async function stopAudioFallback() {
  if (!hasAudioFallbackApis() || audioFallbackTabId === null) return;

  audioFallbackTabId = null;
  try {
    await chrome.runtime.sendMessage({ type: "meetingCopilot:offscreen:stop" });
  } catch (error) {
    // offscreen já pode ter sido descartado
  }
  console.log(`${LOG_PREFIX} fallback de áudio parado`);
}

async function ingestWhisperSegment({ text, meetingId }) {
  const trimmed = String(text ?? "").trim();
  if (!trimmed) return;

  const now = new Date().toISOString();
  const item = {
    id: `whisper-${Date.now()}-${Math.round(Math.random() * 1e6)}`,
    meetingId: meetingId ?? null,
    time: now,
    startedAt: now,
    finalizedAt: now,
    speaker: "Áudio (fallback)",
    text: trimmed,
    source: "whisper",
    url: null
  };

  await enqueueStorageOperation("ao salvar segmento do Whisper", () => appendTranscript(item));
  runAnalysis?.("segmento whisper").catch(() => {});
}

function handleAudioMessage(message, sender) {
  if (message?.type === "meetingCopilot:captionsUnavailable") {
    const tabId = sender?.tab?.id;
    if (typeof tabId !== "number") return Promise.resolve({ ok: false });
    return startAudioFallback(tabId).then((started) => ({ ok: started }));
  }

  if (message?.type === "meetingCopilot:captionsRecovered") {
    return stopAudioFallback().then(() => ({ ok: true }));
  }

  if (message?.type === "meetingCopilot:whisperSegment") {
    return ingestWhisperSegment(message).then(() => ({ ok: true }));
  }

  return undefined;
}

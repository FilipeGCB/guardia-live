(function bootstrapMeetingCopilot() {
  "use strict";

  const LOG_PREFIX = "[MeetingCopilot]";
  const HEARTBEAT_MS = 5000;
  const MAX_SEND_ATTEMPTS = 5;
  const STORAGE_RETRY_DELAY_MS = 500;
  const MAX_STORAGE_RETRY_DELAY_MS = 5000;
  const INVALIDATION_PATTERNS = [
    "Extension context invalidated",
    "Extension context was invalidated",
    "chrome.runtime is undefined",
    "reading 'sendMessage'",
    "reading 'runtime'"
  ];

  const ext = globalThis.browser ?? globalThis.chrome;
  const scheduleRetry = globalThis.setTimeout?.bind(globalThis) || null;
  const cancelRetry = globalThis.clearTimeout?.bind(globalThis) || null;
  const scheduleInterval = globalThis.setInterval?.bind(globalThis) || null;
  const cancelInterval = globalThis.clearInterval?.bind(globalThis) || null;

  function infoLog(event, details) {
    globalThis.console?.log?.(`${LOG_PREFIX} ${event}`, details);
  }

  function debugLog(event, details) {
    if (globalThis.__MEETING_COPILOT_DIAGNOSTICS__ !== true) return;
    globalThis.console?.log?.(`${LOG_PREFIX} ${event}`, details);
  }

  if (typeof globalThis.__MEETING_COPILOT_DIAGNOSTICS__ !== "boolean") {
    globalThis.__MEETING_COPILOT_DIAGNOSTICS__ = false;
  }

  // Uma reinjeção (reload da extensão, sweep do service worker) executa este
  // arquivo de novo no MESMO isolated world da instância anterior. Matar a
  // anterior aqui é o que impede observer duplicado e content script zumbi.
  // `flush: false` é essencial aqui: a instância nova reassume o mesmo elemento
  // de legenda e acumula até o debounce real. Se a antiga gravasse o parcial,
  // a fala apareceria duas vezes — truncada e completa, com ids diferentes.
  const previousTeardown = globalThis.__MEETING_COPILOT_TEARDOWN__;
  if (typeof previousTeardown === "function") {
    try {
      previousTeardown("reinjeção", { flush: false });
    } catch (error) {
      globalThis.console?.warn?.(
        `${LOG_PREFIX} instância anterior não pôde ser encerrada`,
        error
      );
    }
  }

  const api = globalThis.MeetingTranscript;
  if (!api?.SegmentAssembler || !api?.TeamsCaptionSource) {
    globalThis.console?.error?.(
      `${LOG_PREFIX} transcript.js não carregou; captura abortada`
    );
    return;
  }

  const { SegmentAssembler, TeamsCaptionSource } = api;
  const assembler = new SegmentAssembler();
  const segmentMeetingIds = new Map();
  const pendingRetries = new Set();
  const candidateMeetingId = globalThis.crypto?.randomUUID?.() ||
    `meeting-${Date.now()}`;

  let storageWriteQueue = Promise.resolve();
  let source = null;
  let currentMeetingId = null;
  let sessionChangeObserved = false;
  let sessionChangeListener = null;
  let heartbeatTimer = null;
  let sourceStarted = false;
  let stopped = false;
  let invalidationReported = false;
  let resolvingSession = false;

  function contextAlive() {
    try {
      return Boolean(ext?.runtime?.id);
    } catch (error) {
      return false;
    }
  }

  function isRunning() {
    return !stopped && contextAlive();
  }

  function isInvalidationError(error) {
    const message = String(error?.message ?? error);
    return INVALIDATION_PATTERNS.some((pattern) => message.includes(pattern));
  }

  function teardown(reason, { flush = true } = {}) {
    if (stopped) return;
    // Marcado antes do flush: se o próprio flush disparar handleStorageError,
    // o teardown não reentra.
    stopped = true;

    if (flush && contextAlive()) {
      try {
        const pending = assembler.flush?.();
        if (pending) saveTranscriptItem(pending, { force: true });
      } catch (error) {
        globalThis.console?.error?.(
          `${LOG_PREFIX} falha ao finalizar segmento pendente`,
          error
        );
      }
    }

    try {
      source?.stop?.();
    } catch (error) {
      // observer já desconectado
    }

    if (heartbeatTimer !== null) {
      cancelInterval?.(heartbeatTimer);
      heartbeatTimer = null;
    }

    for (const retry of pendingRetries) {
      cancelRetry?.(retry.timer);
      retry.resolve();
    }
    pendingRetries.clear();
    segmentMeetingIds.clear();

    if (sessionChangeListener) {
      try {
        ext?.storage?.onChanged?.removeListener?.(sessionChangeListener);
      } catch (error) {
        // contexto já invalidado: o listener morre junto
      }
      sessionChangeListener = null;
    }

    if (globalThis.__MEETING_COPILOT_TEARDOWN__ === teardown) {
      globalThis.__MEETING_COPILOT_TEARDOWN__ = null;
    }
    if (globalThis.__MEETING_COPILOT_PROBE__ === probe) {
      globalThis.__MEETING_COPILOT_PROBE__ = null;
    }
    globalThis.__MEETING_COPILOT_ALIVE__ = false;

    infoLog("source parado", { reason });
  }

  function probe() {
    return !stopped && contextAlive();
  }

  function reportInvalidation(operation) {
    if (stopped) return;

    if (!invalidationReported) {
      invalidationReported = true;
      globalThis.console?.warn?.(
        `⚠️ Contexto da extensão invalidado; ${operation} interrompida. ` +
        "O service worker reinjeta a captura automaticamente."
      );
    }

    teardown("contexto invalidado", { flush: false });
  }

  function handleStorageError(error, operation) {
    if (!contextAlive() || isInvalidationError(error)) {
      reportInvalidation(operation);
      return;
    }

    globalThis.console?.error?.(`❌ Falha ${operation}`, error);
  }

  function storageRetryDelay(attempt) {
    return Math.min(
      STORAGE_RETRY_DELAY_MS * (2 ** Math.min(attempt - 1, 4)),
      MAX_STORAGE_RETRY_DELAY_MS
    );
  }

  function waitBeforeStorageRetry(attempt) {
    return new Promise((resolve) => {
      if (!scheduleRetry) {
        resolve();
        return;
      }

      const retry = { timer: null, resolve };
      retry.timer = scheduleRetry(() => {
        pendingRetries.delete(retry);
        resolve();
      }, storageRetryDelay(attempt));
      pendingRetries.add(retry);
    });
  }

  // `force` é o caminho do teardown: a instância já está morrendo, mas o
  // contexto ainda vale, então vale a pena tentar gravar o que está em voo.
  async function sendMessageWithRetry(message, operation, { force = false } = {}) {
    const canSend = () => (force ? contextAlive() : isRunning());
    const maxAttempts = force ? 1 : MAX_SEND_ATTEMPTS;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      if (!canSend()) return null;

      try {
        const response = await ext.runtime.sendMessage(message);
        if (!response?.ok) {
          throw new Error(response?.error || "Falha desconhecida no storage");
        }
        return response;
      } catch (error) {
        handleStorageError(error, operation);
        if (!canSend()) return null;

        if (attempt >= maxAttempts) {
          if (!force) {
            globalThis.console?.error?.(
              `❌ ${operation}: desisti após ${attempt} tentativas`,
              { type: message.type, id: message.item?.id || null }
            );
          }
          return null;
        }

        debugLog("storage retry agendado", {
          type: message.type,
          id: message.item?.id || null,
          attempt,
          delayMs: storageRetryDelay(attempt)
        });
        await waitBeforeStorageRetry(attempt);
      }
    }

    return null;
  }

  function saveTranscriptItem(segment, { force = false } = {}) {
    if (force ? !contextAlive() : !isRunning()) {
      debugLog("saveTranscriptItem ignorado", { id: segment.id });
      return;
    }

    const item = {
      id: segment.id,
      meetingId: segmentMeetingIds.get(segment.id) ?? currentMeetingId,
      time: segment.finalizedAt,
      startedAt: segment.startedAt,
      finalizedAt: segment.finalizedAt,
      speaker: segment.speaker,
      text: segment.text,
      url: location.href
    };
    segmentMeetingIds.delete(segment.id);

    infoLog("segmento finalizado", {
      id: item.id,
      speaker: item.speaker,
      text: item.text,
      finalizedAt: item.finalizedAt
    });

    const operation = storageWriteQueue.then(async () => {
      if (force ? !contextAlive() : !isRunning()) return;

      const response = await sendMessageWithRetry({
        type: "meetingCopilot:appendTranscript",
        item
      }, "ao salvar transcript", { force });
      if (!response) return;

      infoLog("storage salvo", {
        id: item.id,
        speaker: item.speaker,
        text: item.text
      });
    });

    // A fila nunca fica presa numa falha: o catch libera o próximo item.
    storageWriteQueue = operation.catch((error) => {
      handleStorageError(error, "ao salvar transcript");
    });
  }

  function handleObservation(observation) {
    if (!isRunning()) return;

    if (!segmentMeetingIds.has(observation.id)) {
      segmentMeetingIds.set(observation.id, currentMeetingId);
    }
    assembler.observe(observation).forEach(saveTranscriptItem);
  }

  function handleFinalized({ id, finalizedAt }) {
    if (!isRunning()) return;

    const segment = assembler.finalize(id, finalizedAt);
    if (segment) saveTranscriptItem(segment);
  }

  function startSource(meetingId) {
    if (!isRunning() || sourceStarted) return;

    sourceStarted = true;
    if (!sessionChangeObserved) currentMeetingId = meetingId ?? null;
    source.start();

    infoLog("source iniciado", {
      url: location.href,
      meetingId: currentMeetingId
    });
  }

  function resolveMeetingSession() {
    if (resolvingSession || !isRunning()) return;
    resolvingSession = true;

    sendMessageWithRetry({
      type: "meetingCopilot:getMeetingSession",
      candidateMeetingId
    }, "ao carregar sessão da reunião").then((response) => {
      resolvingSession = false;
      if (!sessionChangeObserved && response?.meetingId) {
        currentMeetingId = response.meetingId;
      }
      // A captura nunca fica refém do storage: se a sessão não resolveu,
      // o meetingId fica nulo e o service worker anexa à reunião atual.
      startSource(response?.meetingId ?? null);
    }).catch((error) => {
      resolvingSession = false;
      handleStorageError(error, "ao carregar sessão da reunião");
      startSource(null);
    });
  }

  function heartbeat() {
    if (!contextAlive()) {
      reportInvalidation("captura");
      return;
    }

    if (currentMeetingId === null) resolveMeetingSession();
  }

  function loadDebugPreference() {
    try {
      const request = ext?.storage?.local?.get?.({ meetingCopilotDebug: false });
      Promise.resolve(request).then((result) => {
        if (result?.meetingCopilotDebug !== true) return;
        globalThis.__MEETING_COPILOT_DIAGNOSTICS__ = true;
        infoLog("diagnóstico detalhado ativado", { fonte: "storage" });
      }).catch(() => {});
    } catch (error) {
      // storage indisponível: segue com diagnóstico desligado
    }
  }

  function logPersistedTranscript() {
    let request;

    try {
      request = ext?.storage?.local?.get?.("meetingTranscript");
    } catch (error) {
      handleStorageError(error, "ao carregar histórico");
      return;
    }

    Promise.resolve(request).then((result) => {
      const transcript = result?.meetingTranscript || [];
      infoLog("histórico persistido", { falas: transcript.length });
      if (globalThis.__MEETING_COPILOT_DIAGNOSTICS__ === true) {
        globalThis.console?.table?.(transcript);
      }
    }).catch((error) => {
      handleStorageError(error, "ao carregar histórico");
    });
  }

  source = new TeamsCaptionSource({
    onObservation: handleObservation,
    onFinalized: handleFinalized,
    // Captions continuam a fonte primária. Isto só avisa o service worker
    // para tentar o fallback de áudio (Whisper) — se o usuário ligou a opção.
    onCaptionsUnavailable: () => {
      if (!isRunning()) return;
      ext.runtime.sendMessage({ type: "meetingCopilot:captionsUnavailable" }).catch(() => {});
    },
    onCaptionsRecovered: () => {
      if (!isRunning()) return;
      ext.runtime.sendMessage({ type: "meetingCopilot:captionsRecovered" }).catch(() => {});
    }
  });

  sessionChangeListener = (changes, area) => {
    if (
      area === "local" &&
      typeof changes.meetingSessionId?.newValue === "string"
    ) {
      sessionChangeObserved = true;
      currentMeetingId = changes.meetingSessionId.newValue;
    }
  };
  ext?.storage?.onChanged?.addListener?.(sessionChangeListener);

  globalThis.__MEETING_COPILOT_TEARDOWN__ = teardown;
  globalThis.__MEETING_COPILOT_PROBE__ = probe;
  globalThis.__MEETING_COPILOT_ALIVE__ = true;

  globalThis.addEventListener?.("pagehide", () => teardown("pagehide"));

  if (scheduleInterval) {
    heartbeatTimer = scheduleInterval(heartbeat, HEARTBEAT_MS);
  }

  infoLog("bootstrap", {
    url: location.href,
    contextoValido: contextAlive()
  });

  loadDebugPreference();
  logPersistedTranscript();
  resolveMeetingSession();
})();

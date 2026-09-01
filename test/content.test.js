const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const { SegmentAssembler } = require("../transcript.js");

const contentSource = fs.readFileSync(
  path.join(__dirname, "..", "content.js"),
  "utf8"
);

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((promiseResolve, promiseReject) => {
    resolve = promiseResolve;
    reject = promiseReject;
  });

  return { promise, resolve, reject };
}

async function settle() {
  for (let index = 0; index < 5; index += 1) {
    await Promise.resolve();
  }

  await new Promise((resolve) => setImmediate(resolve));
}

function createStorage() {
  const initialRead = deferred();
  const messageCalls = [];
  const calls = [];
  const storageChangeListeners = [];

  const sessionResponses = [];

  const storage = {
    local: {
      get(key) {
        if (key === "meetingTranscript") {
          calls.push({ type: "initialGet", key });
          return initialRead.promise;
        }

        if (key && Object.hasOwn(key, "meetingCopilotDebug")) {
          return Promise.resolve({ meetingCopilotDebug: false });
        }

        throw new Error(`unexpected storage read: ${JSON.stringify(key)}`);
      }
    },
    onChanged: {
      addListener(listener) {
        storageChangeListeners.push(listener);
      },
      removeListener(listener) {
        const index = storageChangeListeners.indexOf(listener);
        if (index >= 0) storageChangeListeners.splice(index, 1);
      }
    }
  };

  const runtime = {
    id: "meeting-copilot-test",
    sendMessage(message) {
      if (message.type === "meetingCopilot:getMeetingSession") {
        if (sessionResponses.length) {
          const next = sessionResponses.shift();
          calls.push({ type: "sendMessage", message });
          return next instanceof Error
            ? Promise.reject(next)
            : Promise.resolve(next);
        }

        return Promise.resolve({ ok: true, meetingId: "session-current" });
      }

      const response = deferred();
      messageCalls.push({ message, ...response });
      calls.push({ type: "sendMessage", message });
      return response.promise;
    }
  };

  return {
    storage,
    runtime,
    initialRead,
    messageCalls,
    calls,
    sessionResponses,
    storageChangeListeners,
    changeMeetingSession(meetingId) {
      for (const listener of storageChangeListeners) {
        listener({
          meetingSessionId: { newValue: meetingId }
        }, "local");
      }
    }
  };
}

function loadContent({ useRealAssembler = false, sessionResponses = [] } = {}) {
  const storageState = createStorage();
  storageState.sessionResponses.push(...sessionResponses);
  const segments = new Map();
  const errors = [];
  const warnings = [];
  const logs = [];
  const retryCallbacks = [];
  const intervals = new Map();
  const pageListeners = new Map();
  const sources = [];
  let nextTimerId = 1;

  class FakeSegmentAssembler {
    observe() {
      return [];
    }

    finalize(id, finalizedAt) {
      const segment = segments.get(id);
      return segment ? { ...segment, finalizedAt } : null;
    }
  }

  class FakeTeamsCaptionSource {
    constructor(callbacks) {
      this.callbacks = callbacks;
      this.startCount = 0;
      this.stopCount = 0;
      sources.push(this);
    }

    start() {
      this.startCount += 1;
    }

    stop() {
      this.stopCount += 1;
    }
  }

  const chromeApi = {
    storage: storageState.storage,
    runtime: storageState.runtime
  };

  const context = vm.createContext({
    browser: undefined,
    chrome: chromeApi,
    console: {
      log(...args) {
        logs.push(args);
      },
      table() {},
      error(...args) {
        errors.push(args);
      },
      warn(...args) {
        warnings.push(args);
      }
    },
    globalThis: undefined,
    location: { href: "https://teams.microsoft.com/v2/" },
    crypto: { randomUUID: () => "session-candidate" },
    setTimeout(callback) {
      retryCallbacks.push(callback);
      return nextTimerId++;
    },
    clearTimeout() {},
    setInterval(callback) {
      const id = nextTimerId++;
      intervals.set(id, callback);
      return id;
    },
    clearInterval(id) {
      intervals.delete(id);
    },
    addEventListener(name, listener) {
      pageListeners.set(name, listener);
    },
    MeetingTranscript: {
      SegmentAssembler: useRealAssembler ? SegmentAssembler : FakeSegmentAssembler,
      TeamsCaptionSource: FakeTeamsCaptionSource
    }
  });

  context.globalThis = context;

  function run() {
    vm.runInContext(contentSource, context, { filename: "content.js" });
  }

  run();

  const state = {
    ...storageState,
    context,
    errors,
    warnings,
    logs,
    sources,
    get source() {
      return sources[sources.length - 1];
    },
    get callbacks() {
      return sources[sources.length - 1].callbacks;
    },
    get storageChangeListenerCount() {
      return storageState.storageChangeListeners.length;
    },
    reinject() {
      run();
    },
    invalidateContext() {
      delete chromeApi.runtime.id;
    },
    runHeartbeat() {
      assert.ok(intervals.size, "heartbeat precisa estar agendado");
      for (const callback of [...intervals.values()]) callback();
    },
    get heartbeatCount() {
      return intervals.size;
    },
    firePageHide() {
      pageListeners.get("pagehide")?.();
    },
    runNextRetry() {
      assert.ok(retryCallbacks.length, "retry precisa estar agendado");
      retryCallbacks.shift()();
    },
    get pendingRetryCount() {
      return retryCallbacks.length;
    },
    emitObservation(observation) {
      state.callbacks.onObservation(observation);
    },
    emitFinalization(id, finalizedAt) {
      state.callbacks.onFinalized({ id, finalizedAt });
    },
    emitFinalized(segment) {
      segments.set(segment.id, segment);
      state.callbacks.onFinalized({
        id: segment.id,
        finalizedAt: segment.finalizedAt
      });
    }
  };

  return state;
}

async function withoutUnhandledRejection(action) {
  const unhandled = [];
  const listener = (reason) => unhandled.push(reason);
  process.on("unhandledRejection", listener);

  try {
    await action();
    await settle();
    return unhandled;
  } finally {
    process.off("unhandledRejection", listener);
  }
}

test("A: leitura inicial invalidada não gera unhandled rejection", () => {
  const contentPath = path.join(__dirname, "..", "content.js");
  const childScript = `
    const fs = require("node:fs");
    const vm = require("node:vm");

    const source = fs.readFileSync(${JSON.stringify(contentPath)}, "utf8");
    let rejectInitialRead;
    const initialRead = new Promise((resolve, reject) => {
      rejectInitialRead = reject;
    });

    class FakeSegmentAssembler {}
    class FakeTeamsCaptionSource {
      constructor() {}
      start() {}
      stop() {}
    }

    const context = vm.createContext({
      browser: undefined,
      chrome: {
        runtime: {
          id: "meeting-copilot-test",
          sendMessage: async () => ({ ok: true, meetingId: "session-test" })
        },
        storage: {
          local: { get: () => initialRead },
          onChanged: { addListener() {}, removeListener() {} }
        }
      },
      console: { log() {}, table() {}, error() {}, warn() {} },
      globalThis: undefined,
      location: { href: "https://teams.microsoft.com/v2/" },
      crypto: { randomUUID: () => "session-candidate" },
      setTimeout,
      clearTimeout,
      setInterval,
      clearInterval,
      addEventListener() {},
      MeetingTranscript: {
        SegmentAssembler: FakeSegmentAssembler,
        TeamsCaptionSource: FakeTeamsCaptionSource
      }
    });

    context.globalThis = context;
    vm.runInContext(source, context, { filename: "content.js" });
    rejectInitialRead(new Error("Extension context invalidated"));
    setTimeout(() => process.exit(0), 100);
  `;

  const result = spawnSync(process.execPath, ["-e", childScript], {
    encoding: "utf8"
  });

  assert.equal(result.status, 0, result.stderr);
});

test("B: falha de persistência por invalidation é tratada", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  const unhandled = await withoutUnhandledRejection(async () => {
    state.emitFinalized({
      id: "segment-1",
      speaker: "João",
      text: "Concordo",
      startedAt: "2026-08-31T12:00:00.000Z",
      finalizedAt: "2026-08-31T12:00:01.000Z"
    });
    await settle();
    state.messageCalls[0].reject(new Error("Extension context invalidated"));
  });

  assert.deepEqual(unhandled, []);
  assert.equal(state.errors.length, 0);
  assert.equal(state.warnings.length, 1);
  assert.match(String(state.warnings[0][0]), /Contexto da extensão invalidado/);
  assert.equal(state.source.stopCount, 1);
});

test("C: depois da invalidation, novas persistências são ignoradas", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitFinalized({
    id: "segment-1",
    speaker: "João",
    text: "Concordo",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();
  state.messageCalls[0].reject(new Error("Extension context invalidated"));
  await settle();

  state.emitFinalized({
    id: "segment-2",
    speaker: "Maria",
    text: "Acho que sim",
    startedAt: "2026-08-31T12:00:02.000Z",
    finalizedAt: "2026-08-31T12:00:03.000Z"
  });
  await settle();

  assert.equal(state.messageCalls.length, 1);
});

test("D: falha comum continua sendo reportada sem invalidar o contexto", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitFinalized({
    id: "segment-1",
    speaker: "João",
    text: "Concordo",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();
  state.messageCalls[0].resolve({
    ok: false,
    error: "Storage quota exceeded"
  });
  await settle();
  state.runNextRetry();
  await settle();

  assert.equal(state.messageCalls.length, 2);
  assert.equal(state.messageCalls[1].message.item.id, "segment-1");
  state.messageCalls[1].resolve({ ok: true });
  await settle();

  state.emitFinalized({
    id: "segment-2",
    speaker: "Maria",
    text: "Acho que sim",
    startedAt: "2026-08-31T12:00:02.000Z",
    finalizedAt: "2026-08-31T12:00:03.000Z"
  });
  await settle();

  assert.equal(state.messageCalls.length, 3);
  assert.equal(state.errors.length, 1);
  assert.match(String(state.errors[0][1]), /Storage quota exceeded/);
  assert.equal(state.source.stopCount, 0);
});

test("E: storageWriteQueue recupera após rejeição do canal", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitFinalized({
    id: "segment-1",
    speaker: "João",
    text: "Primeira fala",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();
  state.messageCalls[0].reject(new Error("Temporary message failure"));
  await settle();
  state.runNextRetry();
  await settle();

  assert.equal(state.messageCalls[1].message.item.id, "segment-1");
  state.messageCalls[1].resolve({ ok: true });
  await settle();

  state.emitFinalized({
    id: "segment-2",
    speaker: "João",
    text: "Segunda fala",
    startedAt: "2026-08-31T12:00:02.000Z",
    finalizedAt: "2026-08-31T12:00:03.000Z"
  });
  await settle();
  state.messageCalls[2].resolve({ ok: true });
  await settle();

  assert.deepEqual(
    state.messageCalls.map(({ message }) => message.item.id),
    ["segment-1", "segment-1", "segment-2"]
  );
});

test("F: novo contexto do Teams cria e inicia uma source independente", async () => {
  const first = loadContent();
  const second = loadContent();
  await settle();

  assert.notEqual(first.source, second.source);
  assert.equal(first.source.startCount, 1);
  assert.equal(second.source.startCount, 1);
});

test("G: observation finalizada por silêncio percorre assembler e chega ao storage", async () => {
  const state = loadContent({ useRealAssembler: true });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitObservation({
    id: "caption-1",
    observedAt: "2026-08-31T12:00:00.000Z",
    speaker: "João",
    text: "Teste de uma frase"
  });
  state.emitFinalization("caption-1", "2026-08-31T12:00:01.500Z");
  await settle();

  const message = state.messageCalls[0].message;
  state.messageCalls[0].resolve({ ok: true });
  await settle();

  assert.equal(message.type, "meetingCopilot:appendTranscript");
  assert.deepEqual(JSON.parse(JSON.stringify(message.item)), {
    id: "caption-1",
    meetingId: "session-current",
    time: "2026-08-31T12:00:01.500Z",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.500Z",
    speaker: "João",
    text: "Teste de uma frase",
    url: "https://teams.microsoft.com/v2/"
  });
});

test("H: fila serial preserva duas falas com mesmo speaker e texto, mas IDs distintos", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  const base = {
    speaker: "João",
    text: "Concordo",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.500Z"
  };
  state.emitFinalized({ id: "segment-1", ...base });
  state.emitFinalized({ id: "segment-2", ...base });
  await settle();

  assert.equal(state.messageCalls.length, 1);
  state.messageCalls[0].resolve({ ok: true });
  await settle();

  assert.equal(state.messageCalls.length, 2);
  state.messageCalls[1].resolve({ ok: true });
  await settle();

  assert.deepEqual(
    state.messageCalls.map(({ message }) => message.item.id),
    ["segment-1", "segment-2"]
  );
});

test("I: segmento ativo preserva a sessão em que começou", async () => {
  const state = loadContent({ useRealAssembler: true });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitObservation({
    id: "caption-old",
    observedAt: "2026-08-31T12:00:00.000Z",
    speaker: "João",
    text: "Fala iniciada antes da troca"
  });
  state.changeMeetingSession("session-new");
  state.emitFinalization("caption-old", "2026-08-31T12:00:02.000Z");
  await settle();

  assert.equal(state.messageCalls[0].message.item.meetingId, "session-current");
});

test("J: segmento iniciado depois da troca usa a sessão nova", async () => {
  const state = loadContent({ useRealAssembler: true });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.changeMeetingSession("session-new");
  state.emitObservation({
    id: "caption-new",
    observedAt: "2026-08-31T12:00:03.000Z",
    speaker: "Maria",
    text: "Fala da reunião nova"
  });
  state.emitFinalization("caption-new", "2026-08-31T12:00:05.000Z");
  await settle();

  assert.equal(state.messageCalls[0].message.item.meetingId, "session-new");
});

test("K: contexto invalidado derruba a instância inteira sem loop", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  assert.equal(state.heartbeatCount, 1);
  state.invalidateContext();
  state.runHeartbeat();

  assert.equal(state.source.stopCount, 1, "observer precisa ser desconectado");
  assert.equal(state.heartbeatCount, 0, "heartbeat precisa parar");
  assert.equal(state.storageChangeListenerCount, 0, "listener precisa sair");
  assert.equal(state.context.__MEETING_COPILOT_ALIVE__, false);
  assert.equal(state.context.__MEETING_COPILOT_TEARDOWN__, null);
  assert.equal(state.warnings.length, 1);
  assert.match(String(state.warnings[0][0]), /Contexto da extensão invalidado/);

  state.emitFinalized({
    id: "segment-zumbi",
    speaker: "João",
    text: "Não deveria salvar",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();

  assert.equal(state.messageCalls.length, 0);
  assert.equal(state.errors.length, 0);
});

test("L: retry é limitado e a fila continua liberada", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitFinalized({
    id: "segment-1",
    speaker: "João",
    text: "Fala que falha sempre",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();

  for (let attempt = 1; attempt <= 5; attempt += 1) {
    assert.equal(state.messageCalls.length, attempt);
    state.messageCalls[attempt - 1].reject(new Error("Falha persistente"));
    await settle();

    if (attempt < 5) {
      state.runNextRetry();
      await settle();
    }
  }

  assert.equal(state.messageCalls.length, 5, "retry precisa ter teto");
  assert.equal(state.pendingRetryCount, 0, "não pode sobrar retry agendado");
  assert.ok(
    state.errors.some(([message]) => /desisti após 5 tentativas/.test(String(message))),
    "desistência precisa ser logada"
  );

  state.emitFinalized({
    id: "segment-2",
    speaker: "Maria",
    text: "Fala seguinte",
    startedAt: "2026-08-31T12:00:02.000Z",
    finalizedAt: "2026-08-31T12:00:03.000Z"
  });
  await settle();

  assert.equal(state.messageCalls.length, 6, "fila não pode ficar bloqueada");
  assert.equal(state.messageCalls[5].message.item.id, "segment-2");
});

test("M: reinjeção encerra a instância anterior e assume a captura", async () => {
  const state = loadContent();
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  const original = state.source;
  state.reinject();
  await settle();

  assert.equal(state.sources.length, 2);
  assert.equal(original.stopCount, 1, "instância antiga precisa morrer");
  assert.notEqual(state.source, original);
  assert.equal(state.source.startCount, 1);
  assert.equal(
    state.storageChangeListenerCount,
    1,
    "não pode acumular listener de storage"
  );

  original.callbacks.onFinalized({
    id: "segment-antigo",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();

  assert.equal(state.messageCalls.length, 0, "instância antiga não persiste");
});

test("N: sessão indisponível não impede a captura", async () => {
  const state = loadContent({
    sessionResponses: Array.from(
      { length: 5 },
      () => new Error("Service worker indisponível")
    )
  });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  for (let attempt = 1; attempt < 5; attempt += 1) {
    state.runNextRetry();
    await settle();
  }

  assert.equal(state.source.startCount, 1, "captura precisa iniciar assim mesmo");

  state.emitFinalized({
    id: "segment-1",
    speaker: "João",
    text: "Fala sem sessão",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.000Z"
  });
  await settle();

  assert.equal(state.messageCalls.length, 1);
  assert.equal(state.messageCalls[0].message.item.meetingId, null);
});

test("O: pagehide finaliza o segmento em andamento antes de morrer", async () => {
  const state = loadContent({ useRealAssembler: true });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitObservation({
    id: "caption-1",
    observedAt: "2026-08-31T12:00:00.000Z",
    speaker: "João",
    text: "Fala interrompida pela navegação"
  });
  state.firePageHide();
  await settle();

  assert.equal(state.messageCalls.length, 1);
  assert.equal(
    state.messageCalls[0].message.item.text,
    "Fala interrompida pela navegação"
  );
  assert.equal(state.source.stopCount, 1);
});

test("P: reinjeção com fala em andamento não parte nem duplica o segmento", async () => {
  const state = loadContent({ useRealAssembler: true });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  // Fala ainda em formação quando a extensão é recarregada.
  state.emitObservation({
    id: "caption-1",
    observedAt: "2026-09-01T12:00:00.000Z",
    speaker: "João",
    text: "Isso é o começo da"
  });
  await settle();
  assert.equal(state.messageCalls.length, 0, "parcial não pode ser persistido");

  state.reinject();
  await settle();

  assert.equal(
    state.messageCalls.length,
    0,
    "reinjeção não pode gravar o parcial truncado"
  );

  // A instância nova reassume e finaliza a fala inteira, uma única vez.
  state.emitObservation({
    id: "caption-2",
    observedAt: "2026-09-01T12:00:01.000Z",
    speaker: "João",
    text: "Isso é o começo da frase completa"
  });
  state.emitFinalization("caption-2", "2026-09-01T12:00:02.500Z");
  await settle();

  assert.equal(state.messageCalls.length, 1);
  assert.equal(
    state.messageCalls[0].message.item.text,
    "Isso é o começo da frase completa"
  );
});

test("Q: pagehide continua gravando o parcial, ao contrário da reinjeção", async () => {
  const state = loadContent({ useRealAssembler: true });
  state.initialRead.resolve({ meetingTranscript: [] });
  await settle();

  state.emitObservation({
    id: "caption-1",
    observedAt: "2026-09-01T12:00:00.000Z",
    speaker: "João",
    text: "Fala perdida se não gravar"
  });
  state.firePageHide();
  await settle();

  assert.equal(state.messageCalls.length, 1);
  assert.equal(
    state.messageCalls[0].message.item.text,
    "Fala perdida se não gravar"
  );
});

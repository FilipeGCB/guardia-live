const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const sidePanelSource = fs.readFileSync(
  path.join(__dirname, "..", "sidepanel.js"),
  "utf8"
);

async function settle() {
  for (let index = 0; index < 5; index += 1) {
    await Promise.resolve();
  }

  await new Promise((resolve) => setImmediate(resolve));
}

class FakeElement {
  constructor() {
    this.textContent = "";
    this.innerHTML = "";
    this.scrollTop = 0;
    this.scrollHeight = 100;
    this.listeners = new Map();
  }

  addEventListener(type, callback) {
    this.listeners.set(type, callback);
  }

  querySelectorAll() {
    return [];
  }

  async click() {
    await this.listeners.get("click")?.();
  }
}

function createStorage(initialState = {}, { getError = null, setError = null } = {}) {
  const state = {
    meetingTranscript: [],
    meetingHistory: [],
    meetingSessionId: "meeting-current",
    ...initialState
  };
  const changesListeners = [];
  const writes = [];

  const runtime = {
    async sendMessage(message) {
      if (setError) throw setError;

      if (message.type === "meetingCopilot:newMeeting") {
        const history = state.meetingHistory;
        if (state.meetingTranscript.length) {
          history.push({
            id: state.meetingSessionId,
            endedAt: message.endedAt,
            transcript: state.meetingTranscript
          });
        }
        state.meetingTranscript = [];
        state.meetingSessionId = message.nextMeetingId;
        writes.push({
          meetingHistory: history,
          meetingTranscript: []
        });
      }

      if (message.type === "meetingCopilot:clearTranscript") {
        state.meetingTranscript = [];
        state.meetingSessionId = message.nextMeetingId;
        writes.push({ meetingTranscript: [] });
      }

      return { ok: true };
    }
  };

  return {
    state,
    writes,
    runtime,
    local: {
      async get(query) {
        if (getError) throw getError;

        if (typeof query === "string") {
          return { [query]: state[query] };
        }

        const result = { ...query };
        for (const key of Object.keys(query)) {
          if (Object.hasOwn(state, key)) result[key] = state[key];
        }
        return result;
      },

      async set(values) {
        if (setError) throw setError;

        const changes = {};
        for (const [key, value] of Object.entries(values)) {
          changes[key] = { oldValue: state[key], newValue: value };
          state[key] = value;
        }
        writes.push(values);
        for (const listener of changesListeners) listener(changes, "local");
      }
    },
    onChanged: {
      addListener(listener) {
        changesListeners.push(listener);
      }
    }
  };
}

async function loadSidePanel({ storage, confirmResult = true } = {}) {
  const elements = {
    transcript: new FakeElement(),
    summary: new FakeElement(),
    newMeeting: new FakeElement(),
    clear: new FakeElement()
  };
  const errors = [];
  const context = vm.createContext({
    chrome: { storage, runtime: storage.runtime },
    confirm: () => confirmResult,
    console: {
      error(...args) {
        errors.push(args);
      },
      log() {}
    },
    crypto: { randomUUID: () => "meeting-new" },
    document: {
      getElementById(id) {
        // sidepanel.js referencia muitos elementos novos (abas, chat,
        // config) que estes testes não exercitam — devolve um stub
        // genérico para qualquer id não explicitamente mapeado acima.
        if (!elements[id]) elements[id] = new FakeElement();
        return elements[id];
      },
      querySelectorAll() {
        return [];
      }
    }
  });

  vm.runInContext(sidePanelSource, context, { filename: "sidepanel.js" });
  await settle();
  return { elements, errors };
}

function transcriptItem(overrides = {}) {
  return {
    time: "2026-08-31T12:00:01.000Z",
    speaker: "João",
    text: "Concordo",
    url: "https://teams.microsoft.com/v2/",
    ...overrides
  };
}

test("Side Panel renderiza formatos legado e atual e preserva ao reabrir", async () => {
  const storage = createStorage({
    meetingTranscript: [
      transcriptItem(),
      transcriptItem({
        id: "segment-2",
        startedAt: "2026-08-31T12:00:02.000Z",
        finalizedAt: "2026-08-31T12:00:03.000Z",
        speaker: "Maria",
        text: "Acho que sim"
      })
    ]
  });

  const firstOpen = await loadSidePanel({ storage });
  assert.equal(firstOpen.elements.summary.textContent, "2 falas nesta reunião");
  assert.match(firstOpen.elements.transcript.innerHTML, /João/);
  assert.match(firstOpen.elements.transcript.innerHTML, /Maria/);
  assert.match(firstOpen.elements.transcript.innerHTML, /Acho que sim/);

  const reopened = await loadSidePanel({ storage });
  assert.equal(reopened.elements.summary.textContent, "2 falas nesta reunião");
  assert.match(reopened.elements.transcript.innerHTML, /Concordo/);
});

test("Side Panel reage a meetingTranscript alterado no storage", async () => {
  const storage = createStorage();
  const panel = await loadSidePanel({ storage });
  assert.equal(panel.elements.summary.textContent, "0 falas nesta reunião");

  await storage.local.set({
    meetingTranscript: [transcriptItem({ text: "Chegou via storage" })]
  });
  await settle();

  assert.equal(panel.elements.summary.textContent, "1 fala nesta reunião");
  assert.match(panel.elements.transcript.innerHTML, /Chegou via storage/);
});

test("Nova reunião arquiva transcript atual sem destruir histórico anterior", async () => {
  const previousMeeting = { id: "meeting-old", transcript: [] };
  const currentTranscript = [transcriptItem({ id: "segment-current" })];
  const storage = createStorage({
    meetingTranscript: currentTranscript,
    meetingHistory: [previousMeeting]
  });
  const panel = await loadSidePanel({ storage });

  await panel.elements.newMeeting.click();
  await settle();

  assert.equal(storage.state.meetingTranscript.length, 0);
  assert.equal(storage.state.meetingHistory.length, 2);
  assert.equal(storage.state.meetingHistory[0], previousMeeting);
  assert.equal(storage.state.meetingHistory[1].id, "meeting-current");
  assert.equal(storage.state.meetingHistory[1].transcript, currentTranscript);
  assert.equal(storage.state.meetingSessionId, "meeting-new");
  assert.equal(panel.elements.summary.textContent, "0 falas nesta reunião");
});

test("Nova reunião vazia preserva meetingHistory sem criar item vazio", async () => {
  const previousMeeting = { id: "meeting-old", transcript: [] };
  const storage = createStorage({ meetingHistory: [previousMeeting] });
  const panel = await loadSidePanel({ storage });

  await panel.elements.newMeeting.click();
  await settle();

  assert.equal(storage.state.meetingHistory.length, 1);
  assert.equal(storage.state.meetingHistory[0], previousMeeting);
  assert.equal(storage.state.meetingTranscript.length, 0);
});

test("Limpar cancelado não altera storage", async () => {
  const currentTranscript = [transcriptItem()];
  const storage = createStorage({ meetingTranscript: currentTranscript });
  const panel = await loadSidePanel({ storage, confirmResult: false });

  await panel.elements.clear.click();
  await settle();

  assert.equal(storage.state.meetingTranscript, currentTranscript);
  assert.equal(storage.writes.length, 0);
});

test("Limpar confirmado remove somente transcript atual", async () => {
  const history = [{ id: "meeting-old", transcript: [transcriptItem()] }];
  const storage = createStorage({
    meetingTranscript: [transcriptItem()],
    meetingHistory: history
  });
  const panel = await loadSidePanel({ storage });

  await panel.elements.clear.click();
  await settle();

  assert.equal(storage.state.meetingTranscript.length, 0);
  assert.equal(storage.state.meetingHistory, history);
  assert.equal(storage.state.meetingSessionId, "meeting-new");
  assert.equal(panel.elements.summary.textContent, "0 falas nesta reunião");
});

test("falha ao renderizar é reportada sem unhandled rejection", async () => {
  const storage = createStorage({}, {
    getError: new Error("Temporary storage read failure")
  });
  const unhandled = [];
  const listener = (reason) => unhandled.push(reason);
  process.on("unhandledRejection", listener);

  try {
    const panel = await loadSidePanel({ storage });
    await settle();

    assert.deepEqual(unhandled, []);
    assert.equal(panel.errors.length, 1);
    assert.match(String(panel.errors[0][1]), /Temporary storage read failure/);
  } finally {
    process.off("unhandledRejection", listener);
  }
});

test("falha ao limpar é reportada sem rejeitar o handler", async () => {
  const storage = createStorage({}, {
    setError: new Error("Temporary storage write failure")
  });
  const panel = await loadSidePanel({ storage });

  await panel.elements.clear.click();
  await settle();

  assert.equal(panel.errors.length, 1);
  assert.match(String(panel.errors[0][1]), /Temporary storage write failure/);
});

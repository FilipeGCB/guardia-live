const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const backgroundSource = fs.readFileSync(
  path.join(__dirname, "..", "background.js"),
  "utf8"
);

function loadBackground(
  initialState = {},
  {
    transientGetFailures = 0,
    transientSetFailures = 0,
    tabs = null,
    probeResults = {},
    injectionFailures = {}
  } = {}
) {
  const state = {
    meetingTranscript: [],
    meetingHistory: [],
    meetingClosedSessions: [],
    meetingSessionId: "session-current",
    ...initialState
  };
  const errors = [];
  const installedListeners = [];
  let listener = null;
  let lastListenerReturn;
  let remainingGetFailures = transientGetFailures;
  let remainingSetFailures = transientSetFailures;
  const chrome = {
    sidePanel: {
      async setPanelBehavior() {}
    },
    runtime: {
      onMessage: {
        addListener(nextListener) {
          listener = nextListener;
        }
      }
    },
    storage: {
      local: {
        async get(query) {
          if (remainingGetFailures > 0) {
            remainingGetFailures -= 1;
            throw new Error("Temporary storage failure");
          }

          const result = { ...query };
          for (const key of Object.keys(query)) {
            if (Object.hasOwn(state, key)) result[key] = state[key];
          }
          return result;
        },
        async set(values) {
          if (remainingSetFailures > 0) {
            remainingSetFailures -= 1;
            throw new Error("Temporary storage set failure");
          }

          Object.assign(state, values);
        },
        async remove(keys) {
          for (const key of Array.isArray(keys) ? keys : [keys]) {
            delete state[key];
          }
        }
      }
    }
  };
  const logs = [];
  const warnings = [];
  const scriptingCalls = [];
  const tabQueries = [];

  if (tabs) {
    chrome.tabs = {
      async query(criteria) {
        tabQueries.push(criteria);
        return tabs;
      },
      onUpdated: { addListener() {} }
    };
    chrome.scripting = {
      async executeScript(details) {
        scriptingCalls.push(details);
        const tabId = details.target.tabId;

        if (details.func) {
          const probe = probeResults[tabId];
          if (probe === "throw") throw new Error("sem acesso à aba");
          return [{ result: probe === true }];
        }

        if (injectionFailures[tabId]) throw new Error("injeção falhou");
        return [{ result: null }];
      }
    };
    chrome.runtime.onInstalled = {
      addListener(listener) {
        installedListeners.push(listener);
      }
    };
    chrome.runtime.onStartup = { addListener() {} };
  }

  const context = vm.createContext({
    chrome,
    console: {
      log(...args) {
        logs.push(args);
      },
      warn(...args) {
        warnings.push(args);
      },
      error(...args) {
        errors.push(args);
      }
    }
  });

  vm.runInContext(backgroundSource, context, { filename: "background.js" });
  return {
    state,
    errors,
    logs,
    warnings,
    chrome,
    scriptingCalls,
    tabQueries,
    installedListeners,
    get injectedTabIds() {
      return scriptingCalls
        .filter(({ files }) => Array.isArray(files))
        .map(({ target }) => target.tabId);
    },
    get probedTabIds() {
      return scriptingCalls
        .filter(({ func }) => typeof func === "function")
        .map(({ target }) => target.tabId);
    },
    get listenerReturn() {
      return lastListenerReturn;
    },
    send(message, sender = {}) {
      assert.ok(listener, "background precisa registrar onMessage");
      return new Promise((resolve, reject) => {
        lastListenerReturn = listener(message, sender, resolve);
        if (lastListenerReturn?.then) {
          lastListenerReturn.then(resolve, reject);
        } else if (lastListenerReturn !== true) {
          resolve(lastListenerReturn);
        }
      });
    }
  };
}

function item(id, text = "Fala", meetingId = "session-current") {
  return {
    id,
    meetingId,
    time: "2026-08-31T12:00:01.500Z",
    startedAt: "2026-08-31T12:00:00.000Z",
    finalizedAt: "2026-08-31T12:00:01.500Z",
    speaker: "João",
    text,
    url: "https://teams.microsoft.com/v2/"
  };
}

test("background serializa append antes de Nova reunião sem perder fala", async () => {
  const background = loadBackground();
  const append = background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-1")
  });
  const newMeeting = background.send({
    type: "meetingCopilot:newMeeting",
    nextMeetingId: "session-new",
    endedAt: "2026-08-31T12:00:02.000Z"
  });

  assert.equal(background.listenerReturn, true);
  assert.equal((await append).ok, true);
  assert.equal((await newMeeting).ok, true);
  assert.equal(background.state.meetingTranscript.length, 0);
  assert.equal(background.state.meetingHistory.length, 1);
  assert.equal(background.state.meetingHistory[0].id, "session-current");
  assert.equal(background.state.meetingHistory[0].transcript[0].id, "segment-1");
  assert.equal(background.state.meetingSessionId, "session-new");
});

test("background serializa append depois de Nova reunião na reunião nova", async () => {
  const background = loadBackground({
    meetingTranscript: [item("segment-old")]
  });
  const newMeeting = background.send({
    type: "meetingCopilot:newMeeting",
    nextMeetingId: "session-new",
    endedAt: "2026-08-31T12:00:02.000Z"
  });
  const append = background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-new", "Fala nova", "session-new")
  });

  await Promise.all([newMeeting, append]);
  assert.deepEqual(
    Array.from(background.state.meetingTranscript, ({ id }) => id),
    ["segment-new"]
  );
  assert.deepEqual(
    Array.from(background.state.meetingHistory[0].transcript, ({ id }) => id),
    ["segment-old"]
  );
});

test("append atrasado da sessão anterior entra no histórico correto", async () => {
  const background = loadBackground();
  await background.send({
    type: "meetingCopilot:newMeeting",
    nextMeetingId: "session-new",
    endedAt: "2026-08-31T12:00:02.000Z"
  });

  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-late", "Fala anterior", "session-current")
  });

  assert.equal(background.state.meetingTranscript.length, 0);
  assert.equal(background.state.meetingHistory[0].transcript[0].id, "segment-late");
});

test("retry idempotente de Nova reunião não encerra a sessão nova", async () => {
  const background = loadBackground({
    meetingTranscript: [item("segment-old")]
  });
  const command = {
    type: "meetingCopilot:newMeeting",
    nextMeetingId: "session-new",
    endedAt: "2026-08-31T12:00:02.000Z"
  };

  await background.send(command);
  await background.send(command);

  assert.equal(background.state.meetingSessionId, "session-new");
  assert.deepEqual(
    Array.from(background.state.meetingClosedSessions, ({ id }) => id),
    ["session-current"]
  );
  assert.equal(background.state.meetingHistory.length, 1);
});

test("background repete falha transitória e preserva append", async () => {
  const background = loadBackground({}, { transientGetFailures: 1 });

  const response = await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-1")
  });

  assert.equal(response.ok, true);
  assert.equal(background.state.meetingTranscript[0].id, "segment-1");
  assert.equal(background.errors.length, 1);
});

test("background repete falha transitória de set sem perder append", async () => {
  const background = loadBackground({}, { transientSetFailures: 1 });

  const response = await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-1")
  });

  assert.equal(response.ok, true);
  assert.equal(background.state.meetingTranscript[0].id, "segment-1");
  assert.equal(background.errors.length, 1);
});

test("background não duplica um lifecycle já persistido", async () => {
  const existing = item("segment-1");
  const background = loadBackground({ meetingTranscript: [existing] });

  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: { ...existing, text: "versão repetida" }
  });

  assert.equal(background.state.meetingTranscript.length, 1);
  assert.equal(background.state.meetingTranscript[0], existing);
});

test("background preserva repetição legítima com IDs diferentes", async () => {
  const background = loadBackground();

  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-1", "Concordo")
  });
  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("segment-2", "Concordo")
  });

  assert.deepEqual(
    Array.from(background.state.meetingTranscript, ({ id }) => id),
    ["segment-1", "segment-2"]
  );
});

test("background Limpar preserva meetingHistory", async () => {
  const history = [{ id: "meeting-old", transcript: [item("archived")] }];
  const background = loadBackground({
    meetingTranscript: [item("current")],
    meetingHistory: history
  });

  assert.equal((await background.send({
    type: "meetingCopilot:clearTranscript",
    nextMeetingId: "session-after-clear"
  })).ok, true);
  assert.equal(background.state.meetingTranscript.length, 0);
  assert.equal(background.state.meetingHistory, history);
  assert.equal(background.state.meetingSessionId, "session-after-clear");
});

test("append atrasado da sessão limpa não reaparece no transcript", async () => {
  const background = loadBackground({
    meetingTranscript: [item("current")]
  });
  await background.send({
    type: "meetingCopilot:clearTranscript",
    nextMeetingId: "session-after-clear"
  });

  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("late", "Fala limpa", "session-current")
  });

  assert.equal(background.state.meetingTranscript.length, 0);
  assert.equal(background.state.meetingHistory.length, 0);
});

test("background cria ou retorna a sessão atual de captura", async () => {
  const background = loadBackground({ meetingSessionId: null });

  const first = await background.send({
    type: "meetingCopilot:getMeetingSession",
    candidateMeetingId: "session-created"
  });
  const second = await background.send({
    type: "meetingCopilot:getMeetingSession",
    candidateMeetingId: "session-ignored"
  });

  assert.equal(first.meetingId, "session-created");
  assert.equal(second.meetingId, "session-created");
  assert.equal(background.state.meetingSessionId, "session-created");
});

test("Nova reunião migra transcript legado mesmo sem meetingSessionId", async () => {
  const legacy = { time: "12:00:00", speaker: "João", text: "Legado" };
  const background = loadBackground({
    meetingSessionId: null,
    meetingTranscript: [legacy]
  });

  await background.send({
    type: "meetingCopilot:newMeeting",
    nextMeetingId: "session-new",
    endedAt: "2026-08-31T12:00:02.000Z"
  });

  assert.equal(background.state.meetingTranscript.length, 0);
  assert.equal(background.state.meetingHistory.length, 1);
  assert.equal(background.state.meetingHistory[0].transcript[0], legacy);
  assert.equal(background.state.meetingSessionId, "session-new");
});

async function drain() {
  for (let index = 0; index < 20; index += 1) {
    await new Promise((resolve) => setImmediate(resolve));
  }
}

test("sweep reinjeta apenas nas abas do Teams sem captura viva", async () => {
  const background = loadBackground({}, {
    tabs: [{ id: 1 }, { id: 2 }, { id: 3 }],
    probeResults: { 1: true, 2: false, 3: "throw" }
  });
  await drain();

  assert.equal(background.tabQueries[0].url, "https://teams.microsoft.com/*");
  assert.deepEqual(background.probedTabIds, [1, 2, 3]);
  assert.deepEqual(
    background.injectedTabIds,
    [2, 3],
    "aba com captura viva não pode ser reinjetada"
  );
  assert.equal(
    background.scriptingCalls
      .find(({ files }) => Array.isArray(files))
      .files.join(","),
    "transcript.js,content.js"
  );
  assert.equal(background.errors.length, 0);
});

test("sweep registra falha de injeção sem derrubar as demais abas", async () => {
  const background = loadBackground({}, {
    tabs: [{ id: 1 }, { id: 2 }],
    probeResults: { 1: false, 2: false },
    injectionFailures: { 1: true }
  });
  await drain();

  assert.deepEqual(background.injectedTabIds, [1, 2]);
  assert.equal(background.errors.length, 1);
  assert.match(String(background.errors[0][0]), /Falha ao reinjetar na aba 1/);
});

test("onInstalled dispara nova varredura de reinjeção", async () => {
  const background = loadBackground({}, {
    tabs: [{ id: 7 }],
    probeResults: { 7: false }
  });
  await drain();
  assert.deepEqual(background.injectedTabIds, [7]);

  assert.equal(background.installedListeners.length, 1);
  background.installedListeners[0]({ reason: "update" });
  await drain();

  assert.deepEqual(background.injectedTabIds, [7, 7]);
  assert.ok(
    background.logs.some(([, details]) => /onInstalled:update/.test(
      String(details?.reason)
    )),
    "a reinjeção precisa registrar o motivo"
  );
});

test("fala sem meetingId é preservada na reunião atual", async () => {
  const background = loadBackground();

  const response = await background.send({
    type: "meetingCopilot:appendTranscript",
    item: { ...item("segment-sem-sessao"), meetingId: null }
  });

  assert.equal(response.ok, true);
  assert.equal(background.state.meetingTranscript.length, 1);
  assert.equal(background.state.meetingTranscript[0].id, "segment-sem-sessao");
  assert.equal(background.warnings.length, 1);
});

test("fala sem meetingId não duplica em retry", async () => {
  const background = loadBackground();
  const orphan = { ...item("segment-sem-sessao"), meetingId: null };

  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: orphan
  });
  await background.send({
    type: "meetingCopilot:appendTranscript",
    item: orphan
  });

  assert.equal(background.state.meetingTranscript.length, 1);
});

test("meetingCopilot:getSettings retorna objeto mesmo sem MeetingLlm carregado", async () => {
  const background = loadBackground();
  const response = await background.send({ type: "meetingCopilot:getSettings" });
  assert.equal(response.ok, true);
  assert.equal(typeof response.settings, "object");
});

test("meetingCopilot:setSettings persiste e é lido de volta", async () => {
  const background = loadBackground();
  await background.send({
    type: "meetingCopilot:setSettings",
    settings: { retentionDays: 30 }
  });
  const response = await background.send({ type: "meetingCopilot:getSettings" });
  assert.equal(response.settings.retentionDays, 30);
});

test("meetingCopilot:dismissInsight remove só o insight pedido", async () => {
  const background = loadBackground({
    meetingInsights: [
      { id: "i1", text: "a" },
      { id: "i2", text: "b" }
    ]
  });
  const response = await background.send({
    type: "meetingCopilot:dismissInsight",
    id: "i1"
  });
  assert.equal(response.ok, true);
  assert.deepEqual(
    background.state.meetingInsights.map((i) => i.id),
    ["i2"]
  );
});

test("meetingCopilot:forgetEverything limpa transcript, histórico, estado e insights", async () => {
  const background = loadBackground({
    meetingTranscript: [item("s1")],
    meetingHistory: [{ id: "old", transcript: [item("s0")] }],
    meetingState: { summary: "x" },
    meetingInsights: [{ id: "i1", text: "a" }]
  });
  const response = await background.send({ type: "meetingCopilot:forgetEverything" });
  assert.equal(response.ok, true);
  assert.equal(background.state.meetingTranscript, undefined);
  assert.equal(background.state.meetingHistory, undefined);
  assert.equal(background.state.meetingState, undefined);
  assert.equal(background.state.meetingInsights, undefined);
});

test("append de transcript não quebra mesmo sem MeetingLlm (análise best-effort)", async () => {
  const background = loadBackground();
  const response = await background.send({
    type: "meetingCopilot:appendTranscript",
    item: item("s1")
  });
  await drain();
  assert.equal(response.ok, true);
  assert.equal(background.errors.length, 0);
});

test("captionsUnavailable sem tabCapture/offscreen não derruba nada (Firefox etc.)", async () => {
  const background = loadBackground();
  const response = await background.send({
    type: "meetingCopilot:captionsUnavailable"
  }, { tab: { id: 7 } });
  assert.equal(response.ok, false);
  assert.equal(background.errors.length, 0);
});

test("whisperSegment vazio é ignorado silenciosamente", async () => {
  const background = loadBackground();
  const response = await background.send({
    type: "meetingCopilot:whisperSegment",
    text: "   ",
    meetingId: "session-current"
  });
  assert.equal(response.ok, true);
  assert.equal(background.state.meetingTranscript.length, 0);
});

test("whisperSegment com texto vira item marcado como source whisper no transcript", async () => {
  const background = loadBackground();
  const response = await background.send({
    type: "meetingCopilot:whisperSegment",
    text: "Falei isso pelo áudio",
    meetingId: "session-current"
  });
  assert.equal(response.ok, true);
  assert.equal(background.state.meetingTranscript.length, 1);
  assert.equal(background.state.meetingTranscript[0].source, "whisper");
  assert.equal(background.state.meetingTranscript[0].speaker, "Áudio (fallback)");
  assert.equal(background.state.meetingTranscript[0].text, "Falei isso pelo áudio");
});

const test = require("node:test");
const assert = require("node:assert/strict");

let transcriptApi;

try {
  transcriptApi = require("../transcript.js");
} catch {
  transcriptApi = null;
}

function getApi() {
  assert.ok(
    transcriptApi,
    "transcript.js deve exportar a camada de captura e montagem"
  );

  return transcriptApi;
}

const CAPTION_SELECTORS = transcriptApi?.CAPTION_SELECTORS || [];

function observation(id, observedAt, speaker, text) {
  return { id, observedAt, speaker, text };
}

class FakeClock {
  constructor(start = "2026-08-31T12:00:00.000Z") {
    this.time = Date.parse(start);
    this.nextTimerId = 1;
    this.timers = new Map();
  }

  now = () => this.time;

  setTimeout = (callback, delay) => {
    const id = this.nextTimerId++;
    this.timers.set(id, { at: this.time + delay, callback });
    return id;
  };

  clearTimeout = (id) => {
    this.timers.delete(id);
  };

  tick(delay) {
    const target = this.time + delay;

    while (true) {
      const next = [...this.timers.entries()]
        .filter(([, timer]) => timer.at <= target)
        .sort(([, left], [, right]) => left.at - right.at)[0];

      if (!next) break;

      const [id, timer] = next;
      this.timers.delete(id);
      this.time = timer.at;
      timer.callback();
    }

    this.time = target;
  }

  iso() {
    return new Date(this.time).toISOString();
  }
}

class UncancellableClock extends FakeClock {
  clearTimeout = () => {};
}

class FakeCaption {
  constructor(
    text,
    speaker,
    {
      tagName = "SPAN",
      dataTid = "closed-caption-text",
      authorSelector = '[data-tid="author"]'
    } = {}
  ) {
    this.innerText = text;
    this.textContent = text;
    this.speaker = speaker;
    this.isConnected = true;
    this.tagName = tagName;
    this.dataTid = dataTid;
    this.authorSelector = authorSelector;
    this.parentElement = null;
  }

  querySelector(selector) {
    if (selector !== this.authorSelector || !this.speaker) {
      return null;
    }

    return { innerText: this.speaker, textContent: this.speaker };
  }

  getAttribute(name) {
    if (name === "data-tid") return this.dataTid;
    return null;
  }
}

class FakeCaptionDocument {
  constructor(captionSelector = '[data-tid="closed-caption-text"]') {
    this.captions = [];
    this.captionSelector = captionSelector;
    this.documentElement = { ownerDocument: this };
    this.body = this.createBody();
  }

  createBody() {
    return { ownerDocument: this };
  }

  replaceBody() {
    this.body = this.createBody();
    return this.body;
  }

  contains(node) {
    return node === this.documentElement || node === this.body;
  }

  querySelectorAll(selector) {
    if (selector === this.captionSelector) {
      return this.captions.filter((caption) => caption.isConnected);
    }

    if (CAPTION_SELECTORS.includes(selector)) return [];
    if (selector === "[data-tid]") return this.captions;
    if (selector === "*") return this.captions;

    assert.fail(`seletor inesperado no fake DOM: ${selector}`);
  }
}

class FakeMutationObserver {
  static instances = [];

  constructor(callback) {
    this.callback = callback;
    this.observedRoot = null;
    this.disconnected = false;
    FakeMutationObserver.instances.push(this);
  }

  observe(root) {
    this.observedRoot = root;
  }

  disconnect() {
    this.disconnected = true;
  }

  trigger(mutationTarget, records = [{ target: mutationTarget }]) {
    if (this.disconnected) return;

    const containsTarget = this.observedRoot?.contains?.(mutationTarget) === true;
    if (mutationTarget === this.observedRoot || containsTarget) {
      this.callback(records);
    }
  }
}

function withFakeMutationObserver(run) {
  const original = globalThis.MutationObserver;
  FakeMutationObserver.instances = [];
  globalThis.MutationObserver = FakeMutationObserver;

  try {
    return run();
  } finally {
    globalThis.MutationObserver = original;
  }
}

function connectCapture({ clock, document, segments }) {
  const { SegmentAssembler, TeamsCaptionSource } = getApi();
  const assembler = new SegmentAssembler();

  const source = new TeamsCaptionSource({
    root: document,
    debounceMs: 1500,
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout,
    onObservation: (item) => {
      segments.push(...assembler.observe(item));
    },
    onFinalized: ({ id, finalizedAt }) => {
      const completed = assembler.finalize(id, finalizedAt);
      if (completed) segments.push(completed);
    }
  });

  return { source, assembler };
}

test("monta uma única fala a partir de captions incrementais", () => {
  const { SegmentAssembler } = getApi();
  const assembler = new SegmentAssembler();

  assert.deepEqual(
    assembler.observe(
      observation("caption-1", "2026-08-31T12:00:00.000Z", "João", "Precisamos")
    ),
    []
  );
  assert.deepEqual(
    assembler.observe(
      observation(
        "caption-1",
        "2026-08-31T12:00:00.400Z",
        "João",
        "Precisamos rever"
      )
    ),
    []
  );
  assert.deepEqual(
    assembler.observe(
      observation(
        "caption-1",
        "2026-08-31T12:00:00.800Z",
        "João",
        "Precisamos rever essa estratégia"
      )
    ),
    []
  );

  assert.deepEqual(
    assembler.finalize("caption-1", "2026-08-31T12:00:02.300Z"),
    {
      id: "caption-1",
      speaker: "João",
      text: "Precisamos rever essa estratégia",
      startedAt: "2026-08-31T12:00:00.000Z",
      finalizedAt: "2026-08-31T12:00:02.300Z"
    }
  );
});

test("ignora observações repetidas e não finaliza o mesmo segmento duas vezes", () => {
  const { SegmentAssembler } = getApi();
  const assembler = new SegmentAssembler();
  const item = observation(
    "caption-1",
    "2026-08-31T12:00:00.000Z",
    "João",
    "Precisamos rever essa estratégia"
  );

  assembler.observe(item);
  assembler.observe({ ...item, observedAt: "2026-08-31T12:00:00.200Z" });
  assembler.observe({ ...item, observedAt: "2026-08-31T12:00:00.400Z" });

  const segment = assembler.finalize("caption-1", "2026-08-31T12:00:01.500Z");

  assert.equal(segment.text, item.text);
  assert.equal(assembler.finalize("caption-1", "2026-08-31T12:00:02.000Z"), null);
});

test("finaliza o speaker anterior ao receber uma nova fala", () => {
  const { SegmentAssembler } = getApi();
  const assembler = new SegmentAssembler();

  assembler.observe(
    observation("caption-1", "2026-08-31T12:00:00.000Z", "João", "Precisamos rever essa estratégia")
  );

  assert.deepEqual(
    assembler.observe(
      observation("caption-2", "2026-08-31T12:00:01.000Z", "Maria", "Concordo")
    ),
    [
      {
        id: "caption-1",
        speaker: "João",
        text: "Precisamos rever essa estratégia",
        startedAt: "2026-08-31T12:00:00.000Z",
        finalizedAt: "2026-08-31T12:00:01.000Z"
      }
    ]
  );

  assert.deepEqual(
    assembler.finalize("caption-2", "2026-08-31T12:00:02.500Z"),
    {
      id: "caption-2",
      speaker: "Maria",
      text: "Concordo",
      startedAt: "2026-08-31T12:00:01.000Z",
      finalizedAt: "2026-08-31T12:00:02.500Z"
    }
  );
});

test("permite uma nova fala do mesmo speaker após a finalização", () => {
  const { SegmentAssembler } = getApi();
  const assembler = new SegmentAssembler();

  assembler.observe(
    observation("caption-1", "2026-08-31T12:00:00.000Z", "João", "Precisamos rever essa estratégia")
  );
  const first = assembler.finalize("caption-1", "2026-08-31T12:00:01.500Z");

  assembler.observe(
    observation("caption-2", "2026-08-31T12:00:03.000Z", "João", "Também precisamos olhar PME")
  );
  const second = assembler.finalize("caption-2", "2026-08-31T12:00:04.500Z");

  assert.equal(first.id, "caption-1");
  assert.equal(second.id, "caption-2");
  assert.notEqual(first.text, second.text);
});

test("ignora texto vazio e ruído de whitespace do DOM", () => {
  const { SegmentAssembler } = getApi();
  const assembler = new SegmentAssembler();

  assert.deepEqual(
    assembler.observe(observation("noise", "2026-08-31T12:00:00.000Z", "João", " \n\t ")),
    []
  );
  assert.equal(assembler.finalize("noise", "2026-08-31T12:00:01.500Z"), null);
});

test("preserva o speaker observado no início da caption", () => {
  const { SegmentAssembler } = getApi();
  const assembler = new SegmentAssembler();

  assembler.observe(
    observation("caption-1", "2026-08-31T12:00:00.000Z", "João", "Precisamos")
  );
  assembler.observe(
    observation("caption-1", "2026-08-31T12:00:00.500Z", "João", "Precisamos rever")
  );

  assert.equal(
    assembler.finalize("caption-1", "2026-08-31T12:00:01.500Z").speaker,
    "João"
  );
});

test("finaliza a fala mesmo quando o elemento DOM é removido antes do debounce", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });
  const caption = new FakeCaption("Precisamos", "João");

  document.captions = [caption];
  source.scan();
  caption.isConnected = false;
  document.captions = [];
  source.scan();

  clock.tick(1500);

  assert.deepEqual(segments, [
    {
      id: segments[0].id,
      speaker: "João",
      text: "Precisamos",
      startedAt: "2026-08-31T12:00:00.000Z",
      finalizedAt: "2026-08-31T12:00:01.500Z"
    }
  ]);
});

test("single utterance finalizes after debounce without any subsequent observation", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });
  const caption = new FakeCaption("Precisamos", "João");

  document.captions = [caption];
  source.scan();

  clock.tick(500);
  caption.innerText = "Precisamos rever";
  caption.textContent = "Precisamos rever";
  source.scan();

  clock.tick(500);
  caption.innerText = "Precisamos rever a estratégia";
  caption.textContent = "Precisamos rever a estratégia";
  source.scan();

  clock.tick(1499);
  assert.equal(segments.length, 0);

  clock.tick(1);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [{ speaker: "João", text: "Precisamos rever a estratégia" }]
  );
});

test("integração finaliza uma única fala no silêncio via onFinalized", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const finalized = [];
  const { SegmentAssembler, TeamsCaptionSource } = getApi();
  const assembler = new SegmentAssembler();
  const source = new TeamsCaptionSource({
    root: document,
    debounceMs: 1500,
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout,
    onObservation: (item) => {
      observations.push(item);
      finalized.push(...assembler.observe(item));
    },
    onFinalized: ({ id, finalizedAt }) => {
      const segment = assembler.finalize(id, finalizedAt);
      if (segment) finalized.push(segment);
    }
  });
  const caption = new FakeCaption("Teste de uma frase", "João");

  document.captions = [caption];
  source.start();
  clock.tick(1501);

  assert.equal(observations.length, 1);
  assert.deepEqual(
    finalized.map(({ speaker, text }) => ({ speaker, text })),
    [{ speaker: "João", text: "Teste de uma frase" }]
  );
});

test("identical repeated observations do not postpone finalization", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Teste", "João")];
  source.scan();

  clock.tick(500);
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Teste", "João")];
  source.scan();

  clock.tick(500);
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Teste", "João")];
  source.scan();

  clock.tick(400);
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Teste", "João")];
  source.scan();

  clock.tick(200);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [{ speaker: "João", text: "Teste" }]
  );
});

test("meaningful text change DOES reset debounce", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Precisamos", "João")];
  source.scan();

  clock.tick(500);
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos", "João")];
  source.scan();

  clock.tick(400);
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever", "João")];
  source.scan();

  clock.tick(300);
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever", "João")];
  source.scan();

  clock.tick(1199);
  assert.equal(segments.length, 0);

  clock.tick(1);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [{ speaker: "João", text: "Precisamos rever" }]
  );
});

test("TESTE C — não duplica caption quando o Teams recria o elemento com o mesmo conteúdo", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Precisamos rever essa estratégia", "João")];
  source.scan();
  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever essa estratégia", "João")];
  source.scan();
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 1);
  assert.equal(segments[0].speaker, "João");
  assert.equal(segments[0].text, "Precisamos rever essa estratégia");
});

test("TESTE D — continua a mesma caption quando o Teams recria o elemento com texto atualizado", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Precisamos", "João")];
  source.scan();
  clock.tick(400);

  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever essa estratégia", "João")];
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 1);
  assert.equal(segments[0].text, "Precisamos rever essa estratégia");
  assert.equal(segments[0].speaker, "João");
});

test("mantém o speaker conhecido quando o nó recriado ainda não tem autor", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Precisamos", "João")];
  source.scan();
  clock.tick(400);

  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever essa estratégia", null)];
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 1);
  assert.equal(segments[0].speaker, "João");
  assert.equal(segments[0].text, "Precisamos rever essa estratégia");
});

test("preserva uma nova utterance com o mesmo conteúdo após a finalização", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Precisamos rever essa estratégia", "João")];
  source.scan();
  clock.tick(1500);

  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever essa estratégia", "João")];
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 2);
});

test("TESTE A — cria nova utterance quando o speaker é igual, mas o texto recriado não continua", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  const first = new FakeCaption("Precisamos rever a estratégia", "João");
  document.captions = [first];
  source.scan();

  first.isConnected = false;
  document.captions = [new FakeCaption("Também precisamos olhar PME", "João")];
  source.scan();
  clock.tick(1500);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Precisamos rever a estratégia" },
      { speaker: "João", text: "Também precisamos olhar PME" }
    ]
  );
});

test("TESTE B — preserva repetição legítima da mesma frase em menos de 10 segundos", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  const first = new FakeCaption("Concordo", "João");
  document.captions = [first];
  source.scan();
  clock.tick(1500);

  first.isConnected = false;
  clock.tick(2500);
  const maria = new FakeCaption("Acho que sim", "Maria");
  document.captions = [maria];
  source.scan();
  clock.tick(1500);

  maria.isConnected = false;
  clock.tick(1500);
  document.captions = [new FakeCaption("Concordo", "João")];
  source.scan();
  clock.tick(1500);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Concordo" },
      { speaker: "Maria", text: "Acho que sim" },
      { speaker: "João", text: "Concordo" }
    ]
  );
});

test("speaker desconhecido não reutiliza um estado apenas por speaker quando o texto não continua", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  const first = new FakeCaption("Precisamos rever a estratégia", null);
  document.captions = [first];
  source.scan();
  first.isConnected = false;
  document.captions = [new FakeCaption("Também precisamos olhar PME", null)];
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 2);
});

test("seleciona entre candidatos do mesmo speaker somente o lifecycle com continuidade textual", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const { TeamsCaptionSource } = getApi();
  const source = new TeamsCaptionSource({
    root: document,
    onObservation: (item) => observations.push(item),
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout
  });
  const first = new FakeCaption("Precisamos", "João");
  const second = new FakeCaption("Também", "João");

  document.captions = [first];
  source.scan();
  clock.tick(100);
  document.captions = [first, second];
  source.scan();

  first.isConnected = false;
  second.isConnected = false;
  document.captions = [new FakeCaption("Precisamos rever", "João")];
  source.scan();

  assert.equal(observations.length, 3);
  assert.equal(observations[2].id, observations[0].id);
});

test("recriação simultânea preserva os lifecycles quando textos têm prefixo ambíguo", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const { TeamsCaptionSource } = getApi();
  const source = new TeamsCaptionSource({
    root: document,
    onObservation: (item) => observations.push(item),
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout
  });
  const short = new FakeCaption("Precisamos", "João");
  const long = new FakeCaption("Precisamos rever", "João");

  document.captions = [short];
  source.scan();
  const shortId = observations[0].id;
  clock.tick(100);
  document.captions = [short, long];
  source.scan();
  const longId = observations[1].id;

  short.isConnected = false;
  long.isConnected = false;
  const recreatedShort = new FakeCaption("Precisamos", "João");
  const recreatedLong = new FakeCaption("Precisamos rever", "João");
  document.captions = [recreatedShort, recreatedLong];
  source.scan();

  assert.equal(source.elementStates.get(recreatedShort).id, shortId);
  assert.equal(source.elementStates.get(recreatedLong).id, longId);
});

test("TESTE E — timer antigo não sobrescreve a utterance nova após remoção e recriação", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });
  const first = new FakeCaption("Precisamos", "João");

  document.captions = [first];
  source.scan();
  clock.tick(500);
  first.innerText = "Precisamos rever";
  first.textContent = "Precisamos rever";
  source.scan();
  clock.tick(400);
  first.innerText = "Precisamos rever a estratégia";
  first.textContent = "Precisamos rever a estratégia";
  source.scan();
  clock.tick(300);

  first.isConnected = false;
  document.captions = [];
  source.scan();
  clock.tick(50);
  document.captions = [new FakeCaption("Precisamos rever a estratégia", "João")];
  source.scan();
  clock.tick(350);

  document.captions = [new FakeCaption("Também precisamos olhar PME", "João")];
  source.scan();
  clock.tick(1500);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Precisamos rever a estratégia" },
      { speaker: "João", text: "Também precisamos olhar PME" }
    ]
  );
});

test("não concatena indefinidamente duas falas consecutivas do mesmo speaker", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Precisamos rever essa estratégia", "João")];
  source.scan();
  clock.tick(1500);

  document.captions[0].isConnected = false;
  document.captions = [new FakeCaption("Também precisamos olhar PME", "João")];
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 2);
  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Precisamos rever essa estratégia" },
      { speaker: "João", text: "Também precisamos olhar PME" }
    ]
  );
});

test("candidatos idênticos na mesma varredura não criam lifecycles duplicados", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  document.captions = [new FakeCaption("Olá, tudo bem?", "Filipe")];
  source.scan();

  document.captions[0].isConnected = false;
  document.captions = [
    new FakeCaption("Olá, tudo bem?", "Filipe"),
    new FakeCaption("Olá, tudo bem?", "Filipe")
  ];
  source.scan();
  clock.tick(1500);

  assert.equal(segments.length, 1);
  assert.equal(segments[0].text, "Olá, tudo bem?");
});

test("um timer antigo não finaliza uma atualização posterior da mesma caption", () => {
  const clock = new UncancellableClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });
  const caption = new FakeCaption("Precisamos", "João");

  document.captions = [caption];
  source.scan();
  clock.tick(1000);

  caption.innerText = "Precisamos rever essa estratégia";
  caption.textContent = "Precisamos rever essa estratégia";
  source.scan();

  clock.tick(500);
  assert.equal(segments.length, 0);

  clock.tick(1000);
  assert.equal(segments.length, 1);
  assert.equal(segments[0].text, "Precisamos rever essa estratégia");
});

test("timers globais são chamados com receiver compatível com o browser", () => {
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  const document = new FakeCaptionDocument();
  const scheduled = [];
  const cancelled = [];

  globalThis.setTimeout = function browserSetTimeout(callback, delay) {
    if (this !== globalThis) throw new TypeError("Illegal invocation");
    scheduled.push({ callback, delay });
    return 42;
  };
  globalThis.clearTimeout = function browserClearTimeout(timer) {
    if (this !== globalThis) throw new TypeError("Illegal invocation");
    cancelled.push(timer);
  };

  try {
    const { TeamsCaptionSource } = getApi();
    const source = new TeamsCaptionSource({ root: document });
    document.captions = [new FakeCaption("Teste de timer", "João")];

    source.scan();
    source.stop();

    assert.equal(scheduled.length, 1);
    assert.equal(scheduled[0].delay, 1500);
    assert.deepEqual(cancelled, [42]);
  } finally {
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  }
});

test("troca de speaker no mesmo nó não contamina a fala anterior", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });
  const caption = new FakeCaption("Precisamos rever essa estratégia", "João");

  document.captions = [caption];
  source.scan();
  caption.innerText = "Concordo";
  caption.textContent = "Concordo";
  caption.speaker = "Maria";
  source.scan();
  clock.tick(1500);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Precisamos rever essa estratégia" },
      { speaker: "Maria", text: "Concordo" }
    ]
  );
});

test("não emite observações para captions vazias do DOM", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const { TeamsCaptionSource } = getApi();
  const source = new TeamsCaptionSource({
    root: document,
    onObservation: (item) => observations.push(item),
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout
  });

  document.captions = [new FakeCaption("  \n  ", "João")];
  source.scan();

  assert.deepEqual(observations, []);
});

test("camada DOM encontra caption com o seletor atual", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const { TeamsCaptionSource } = getApi();
  const caption = new FakeCaption("Legenda encontrada", "João");
  const source = new TeamsCaptionSource({
    root: document,
    onObservation: (item) => observations.push(item),
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout
  });

  document.captions = [caption];
  source.scan();

  assert.equal(observations.length, 1);
  assert.equal(observations[0].text, "Legenda encontrada");
});

test("start captura caption que já existia antes da inicialização", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const observations = [];
    const { TeamsCaptionSource } = getApi();
    const source = new TeamsCaptionSource({
      root: document,
      onObservation: (item) => observations.push(item),
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout
    });

    document.captions = [new FakeCaption("Já estava visível", "Maria")];
    source.start();

    assert.equal(observations.length, 1);
    assert.equal(observations[0].text, "Já estava visível");
  });
});

test("mutation dispara scan para caption adicionada depois do start", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const observations = [];
    const { TeamsCaptionSource } = getApi();
    const source = new TeamsCaptionSource({
      root: document,
      onObservation: (item) => observations.push(item),
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout
    });

    source.start();
    document.captions = [new FakeCaption("Chegou depois", "João")];
    FakeMutationObserver.instances[0].trigger(document.body);

    assert.equal(observations.length, 1);
    assert.equal(observations[0].text, "Chegou depois");
  });
});

test("selector zero diagnostica nó textual adicionado por mutation posterior", () => {
  withFakeMutationObserver(() => {
    const originalDiagnostics = globalThis.__MEETING_COPILOT_DIAGNOSTICS__;
    const originalLog = globalThis.console.log;
    const logs = [];
    globalThis.__MEETING_COPILOT_DIAGNOSTICS__ = true;
    globalThis.console.log = (message, details) => logs.push({ message, details });

    try {
      const clock = new FakeClock();
      const document = new FakeCaptionDocument();
      const { TeamsCaptionSource } = getApi();
      const source = new TeamsCaptionSource({
        root: document,
        now: clock.now,
        setTimeout: clock.setTimeout,
        clearTimeout: clock.clearTimeout
      });
      const addedNode = {
        nodeType: 1,
        tagName: "DIV",
        innerText: "Teste de uma frase",
        textContent: "Teste de uma frase",
        getAttribute() {
          return null;
        }
      };

      source.start();
      FakeMutationObserver.instances[0].trigger(document.body, [
        {
          type: "childList",
          target: document.body,
          addedNodes: [addedNode],
          removedNodes: []
        }
      ]);

      const diagnostic = logs.find(({ message }) =>
        message === "[MeetingCopilot][CAPTURE] diagnóstico mutation com seletor zero"
      );
      assert.ok(diagnostic);
      assert.equal(diagnostic.details.candidates.length, 1);
      assert.equal(diagnostic.details.candidates[0].text, "Teste de uma frase");
    } finally {
      globalThis.__MEETING_COPILOT_DIAGNOSTICS__ = originalDiagnostics;
      globalThis.console.log = originalLog;
    }
  });
});

test("start repetido não instala MutationObserver duplicado", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const { TeamsCaptionSource } = getApi();
    const source = new TeamsCaptionSource({
      root: document,
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout
    });

    source.start();
    source.start();

    assert.equal(FakeMutationObserver.instances.length, 1);
  });
});

test("observer continua capturando depois que o root DOM observado é substituído", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const observations = [];
    const { TeamsCaptionSource } = getApi();
    const source = new TeamsCaptionSource({
      root: document,
      onObservation: (item) => observations.push(item),
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout
    });

    source.start();
    document.replaceBody();
    document.captions = [new FakeCaption("Após troca do body", "Maria")];
    FakeMutationObserver.instances[0].trigger(document.documentElement);

    assert.equal(observations.length, 1);
    assert.equal(observations[0].text, "Após troca do body");
  });
});

test("observer continua capturando após navegação SPA no mesmo document", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const observations = [];
    const { TeamsCaptionSource } = getApi();
    const source = new TeamsCaptionSource({
      root: document,
      onObservation: (item) => observations.push(item),
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout
    });

    source.start();
    document.captions = [new FakeCaption("Após navegação SPA", "João")];
    FakeMutationObserver.instances[0].trigger(document.body);

    assert.equal(observations.length, 1);
    assert.equal(observations[0].text, "Após navegação SPA");
  });
});

test("captions off e on retomam captura sem perder segmentos anteriores", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const segments = [];
  const { source } = connectCapture({ clock, document, segments });

  const first = new FakeCaption("Antes de desligar", "João");
  document.captions = [first];
  source.scan();
  first.isConnected = false;
  document.captions = [];
  source.scan();
  clock.tick(1500);

  document.captions = [new FakeCaption("Depois de religar", "Maria")];
  source.scan();
  clock.tick(1500);

  assert.deepEqual(
    segments.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Antes de desligar" },
      { speaker: "Maria", text: "Depois de religar" }
    ]
  );
});

test("lê texto tanto no próprio nó quanto agregado de elemento filho", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const { TeamsCaptionSource } = getApi();
  const ownText = new FakeCaption("Texto no próprio nó", "João");
  const childText = new FakeCaption("", "Maria");
  childText.textContent = "Texto vindo do filho";
  const source = new TeamsCaptionSource({
    root: document,
    onObservation: (item) => observations.push(item),
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout
  });

  document.captions = [ownText, childText];
  source.scan();

  assert.deepEqual(
    observations.map(({ speaker, text }) => ({ speaker, text })),
    [
      { speaker: "João", text: "Texto no próprio nó" },
      { speaker: "Maria", text: "Texto vindo do filho" }
    ]
  );
});

test("encontra speaker em sibling por meio do ancestor comum", () => {
  const clock = new FakeClock();
  const document = new FakeCaptionDocument();
  const observations = [];
  const { TeamsCaptionSource } = getApi();
  const caption = new FakeCaption("Texto da legenda", null);
  caption.parentElement = {
    parentElement: null,
    querySelector(selector) {
      if (selector !== '[data-tid="author"]') return null;
      return { innerText: "Speaker sibling", textContent: "Speaker sibling" };
    }
  };
  const source = new TeamsCaptionSource({
    root: document,
    onObservation: (item) => observations.push(item),
    now: clock.now,
    setTimeout: clock.setTimeout,
    clearTimeout: clock.clearTimeout
  });

  document.captions = [caption];
  source.scan();

  assert.equal(observations.length, 1);
  assert.equal(observations[0].speaker, "Speaker sibling");
});

test("adaptador resolve seletor alternativo de caption do Teams", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument('[data-tid="closed-caption-message"]');
    const segments = [];
    const { source } = connectCapture({ clock, document, segments });

    document.captions = [
      new FakeCaption("Fala em DOM novo", "João", {
        dataTid: "closed-caption-message"
      })
    ];

    source.start();
    clock.tick(1500);

    assert.equal(source.captionSelector, '[data-tid="closed-caption-message"]');
    assert.equal(segments.length, 1);
    assert.equal(segments[0].speaker, "João");
    assert.equal(segments[0].text, "Fala em DOM novo");
  });
});

test("adaptador resolve seletor alternativo de autor", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const segments = [];
    const { source } = connectCapture({ clock, document, segments });

    document.captions = [
      new FakeCaption("Fala com autor alternativo", "Maria", {
        authorSelector: '[data-tid="closed-caption-author"]'
      })
    ];

    source.start();
    clock.tick(1500);

    assert.equal(segments.length, 1);
    assert.equal(segments[0].speaker, "Maria");
  });
});

test("seletor que deixa de casar é reresolvido sem perder captura", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const segments = [];
    const { source } = connectCapture({ clock, document, segments });

    document.captions = [new FakeCaption("Primeira fala", "João")];
    source.start();
    clock.tick(1500);
    assert.equal(source.captionSelector, '[data-tid="closed-caption-text"]');

    // Teams troca o DOM das legendas no meio da reunião.
    document.captionSelector = '[data-tid="caption-text"]';
    document.captions = [
      new FakeCaption("Fala depois da troca", "Maria", {
        dataTid: "caption-text"
      })
    ];
    FakeMutationObserver.instances[0].trigger(document.body);
    clock.tick(1500);

    assert.equal(source.captionSelector, '[data-tid="caption-text"]');
    assert.deepEqual(
      segments.map(({ text }) => text),
      ["Primeira fala", "Fala depois da troca"]
    );
  });
});

test("troca rápida de speaker no mesmo elemento gera dois segmentos", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const segments = [];
    const { source } = connectCapture({ clock, document, segments });

    const caption = new FakeCaption("Bom dia", "João");
    document.captions = [caption];
    source.start();

    caption.innerText = "Bom dia a todos";
    caption.textContent = "Bom dia a todos";
    FakeMutationObserver.instances[0].trigger(document.body);

    caption.innerText = "Tudo certo";
    caption.textContent = "Tudo certo";
    caption.speaker = "Maria";
    FakeMutationObserver.instances[0].trigger(document.body);
    clock.tick(1500);

    assert.deepEqual(
      segments.map(({ speaker, text }) => `${speaker}: ${text}`),
      ["João: Bom dia a todos", "Maria: Tudo certo"]
    );
  });
});

test("diagnóstico de seletor zero repete espaçado sem virar loop", () => {
  withFakeMutationObserver(() => {
    const originalLog = globalThis.console.log;
    const logs = [];
    globalThis.console.log = (message, details) => logs.push({ message, details });

    try {
      const clock = new FakeClock();
      const document = new FakeCaptionDocument();
      const { TeamsCaptionSource } = getApi();
      const source = new TeamsCaptionSource({
        root: document,
        now: clock.now,
        setTimeout: clock.setTimeout,
        clearTimeout: clock.clearTimeout
      });

      source.start();
      source.scan();
      source.scan();

      const zeroLogs = () => logs.filter(({ message }) =>
        message === "[MeetingCopilot] nenhum seletor de caption casou"
      );
      assert.equal(zeroLogs().length, 1, "não pode repetir no mesmo instante");

      clock.tick(30000);
      source.scan();
      assert.equal(zeroLogs().length, 2);

      for (let index = 0; index < 10; index += 1) {
        clock.tick(30000);
        source.scan();
      }
      assert.equal(zeroLogs().length, 5, "diagnóstico precisa ter teto");
    } finally {
      globalThis.console.log = originalLog;
    }
  });
});

test("onCaptionsUnavailable dispara após esgotar as tentativas de diagnóstico", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    const { TeamsCaptionSource } = getApi();
    const unavailable = [];
    const recovered = [];
    const source = new TeamsCaptionSource({
      root: document,
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout,
      onCaptionsUnavailable: () => unavailable.push(clock.iso()),
      onCaptionsRecovered: () => recovered.push(clock.iso())
    });

    source.start();
    // 5 tentativas (MAX_ZERO_DIAGNOSTICS), espaçadas pelo intervalo mínimo.
    for (let i = 0; i < 5; i += 1) {
      source.scan();
      clock.tick(30000);
    }

    assert.equal(unavailable.length, 1);
    assert.equal(recovered.length, 0);

    document.captions = [new FakeCaption("Voltou", "João")];
    source.scan();

    assert.equal(recovered.length, 1);
  });
});

test("onCaptionsUnavailable não dispara enquanto captions aparecem normalmente", () => {
  withFakeMutationObserver(() => {
    const clock = new FakeClock();
    const document = new FakeCaptionDocument();
    document.captions = [new FakeCaption("Fala normal", "João")];
    const { TeamsCaptionSource } = getApi();
    const unavailable = [];
    const source = new TeamsCaptionSource({
      root: document,
      now: clock.now,
      setTimeout: clock.setTimeout,
      clearTimeout: clock.clearTimeout,
      onCaptionsUnavailable: () => unavailable.push(true)
    });

    source.start();
    for (let i = 0; i < 5; i += 1) {
      source.scan();
      clock.tick(30000);
    }

    assert.equal(unavailable.length, 0);
  });
});

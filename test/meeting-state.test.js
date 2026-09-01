const assert = require("node:assert/strict");
const test = require("node:test");
const {
  parseState,
  parseInsights,
  shouldAnalyze,
  extractJson,
  buildStatePrompt,
  buildInsightPrompt,
  SHORTCUTS
} = require("../meeting-state.js");

const transcript = [
  { id: "s1", speaker: "João", text: "Vamos adiar o lançamento para março" },
  { id: "s2", speaker: "Maria", text: "Quem fica responsável pelo relatório?" }
];

test("parseState extrai JSON de resposta cercada em markdown", () => {
  const raw = "Aqui está:\n```json\n" + JSON.stringify({
    summary: "Discussão sobre o cronograma.",
    decisions: [{ text: "Adiar lançamento para março", evidence: ["s1"] }],
    openQuestions: [{ text: "Quem fica com o relatório?", evidence: ["s2"] }]
  }) + "\n```";

  const state = parseState(raw, { transcript, segmentCount: 2, now: "2026-09-01T00:00:00Z" });

  assert.equal(state.summary, "Discussão sobre o cronograma.");
  assert.equal(state.decisions.length, 1);
  assert.deepEqual(state.decisions[0].evidence, ["s1"]);
  assert.equal(state.openQuestions[0].text, "Quem fica com o relatório?");
  assert.equal(state.segmentCount, 2);
});

test("parseState descarta evidência que não existe no transcript", () => {
  const raw = JSON.stringify({
    summary: "x",
    risks: [{ text: "Risco inventado", evidence: ["s1", "s999"] }]
  });

  const state = parseState(raw, { transcript, segmentCount: 2 });

  assert.deepEqual(state.risks[0].evidence, ["s1"]);
});

test("parseState retorna null para lixo sem JSON", () => {
  assert.equal(parseState("não consigo responder agora", { transcript }), null);
});

test("parseInsights limita a MAX_INSIGHTS_PER_BLOCK e deduplica contra já vistos", () => {
  const raw = JSON.stringify({
    insights: [
      { kind: "pergunta", text: "Já visto antes", evidence: ["s1"] },
      { kind: "ponto", text: "Ponto novo A", evidence: ["s2"] },
      { kind: "risco", text: "Ponto novo B", evidence: ["s1"] },
      { kind: "ponto", text: "Ponto novo C (deveria ser cortado)", evidence: [] }
    ]
  });

  const insights = parseInsights(raw, {
    transcript,
    seen: ["Já visto antes"]
  });

  assert.equal(insights.length, 2);
  assert.deepEqual(insights.map((i) => i.text), ["Ponto novo A", "Ponto novo B"]);
});

test("parseInsights aceita {insights: []} como silêncio válido", () => {
  assert.deepEqual(parseInsights('{"insights": []}', { transcript }), []);
});

test("shouldAnalyze só dispara em blocos, não a cada fala", () => {
  assert.equal(shouldAnalyze({ transcriptLength: 3, analyzedCount: 0, blockSize: 8 }), false);
  assert.equal(shouldAnalyze({ transcriptLength: 8, analyzedCount: 0, blockSize: 8 }), true);
  assert.equal(shouldAnalyze({ transcriptLength: 10, analyzedCount: 8, blockSize: 8 }), false);
  assert.equal(shouldAnalyze({ transcriptLength: 16, analyzedCount: 8, blockSize: 8 }), true);
});

test("shouldAnalyze reprocessa se o transcript encolheu (Nova reunião)", () => {
  assert.equal(shouldAnalyze({ transcriptLength: 2, analyzedCount: 10, blockSize: 8 }), true);
});

test("buildStatePrompt inclui o resumo anterior para atualização incremental", () => {
  const prompt = buildStatePrompt(transcript, { summary: "Resumo prévio" });
  assert.match(prompt.messages[0].content, /Resumo prévio/);
  assert.match(prompt.messages[0].content, /João/);
});

test("buildInsightPrompt não inclui contexto quando não há estado prévio", () => {
  const prompt = buildInsightPrompt(transcript, null);
  assert.doesNotMatch(prompt.messages[0].content, /Contexto da reunião/);
});

test("extractJson aceita objeto cru sem cerca de código", () => {
  assert.deepEqual(extractJson('{"a": 1}'), { a: 1 });
});

test("SHORTCUTS cobre os sete atalhos pedidos", () => {
  const ids = SHORTCUTS.map((s) => s.id).sort();
  assert.deepEqual(ids, [
    "contestar", "explicar", "minha-vez", "perguntar",
    "responder", "resumo", "riscos"
  ]);
});

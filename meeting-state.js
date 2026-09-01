(function (root, factory) {
  const api = factory();

  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.MeetingState = api;
  }
})(globalThis, function () {
  const MAX_TRANSCRIPT_CHARS = 24000;
  const MAX_INSIGHTS_PER_BLOCK = 2;
  const EMPTY_STATE = {
    summary: "",
    topics: [],
    decisions: [],
    risks: [],
    openQuestions: [],
    commitments: [],
    segmentCount: 0,
    updatedAt: null
  };

  function emptyState() {
    return JSON.parse(JSON.stringify(EMPTY_STATE));
  }

  function normalizeText(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
  }

  function dedupeKey(value) {
    return normalizeText(value).toLowerCase();
  }

  // O modelo costuma embrulhar o JSON em prosa ou em cerca de código.
  function extractJson(raw) {
    const text = String(raw ?? "");
    const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
    const candidates = [];

    if (fenced) candidates.push(fenced[1]);
    candidates.push(text);

    const start = text.indexOf("{");
    const end = text.lastIndexOf("}");
    if (start >= 0 && end > start) candidates.push(text.slice(start, end + 1));

    const arrayStart = text.indexOf("[");
    const arrayEnd = text.lastIndexOf("]");
    if (arrayStart >= 0 && arrayEnd > arrayStart) {
      candidates.push(text.slice(arrayStart, arrayEnd + 1));
    }

    for (const candidate of candidates) {
      try {
        const parsed = JSON.parse(candidate.trim());
        if (parsed && typeof parsed === "object") return parsed;
      } catch (error) {
        // tenta o próximo candidato
      }
    }

    return null;
  }

  // Só aceita evidências que existem de fato no transcript: o transcript é a
  // fonte factual, o modelo não pode inventar a fala que sustenta a inferência.
  function normalizeEvidence(value, knownIds) {
    const list = Array.isArray(value) ? value : [value];

    return [...new Set(
      list
        .map((entry) => normalizeText(entry?.id ?? entry))
        .filter((id) => id && (!knownIds || knownIds.has(id)))
    )];
  }

  function normalizeItems(value, knownIds, { fields = ["text"] } = {}) {
    if (!Array.isArray(value)) return [];

    const seen = new Set();
    const items = [];

    for (const entry of value) {
      const item = {};
      let primary = "";

      for (const field of fields) {
        const raw = typeof entry === "string" && field === fields[0]
          ? entry
          : entry?.[field];
        item[field] = normalizeText(raw);
        if (field === fields[0]) primary = item[field];
      }

      if (!primary) continue;

      const key = dedupeKey(fields.map((field) => item[field]).join("|"));
      if (seen.has(key)) continue;
      seen.add(key);

      item.evidence = normalizeEvidence(entry?.evidence, knownIds);
      items.push(item);
    }

    return items;
  }

  function parseState(raw, { transcript = [], segmentCount = 0, now = null } = {}) {
    const parsed = extractJson(raw);
    if (!parsed) return null;

    const knownIds = new Set(
      transcript.map((item) => normalizeText(item?.id)).filter(Boolean)
    );

    return {
      summary: normalizeText(parsed.summary || parsed.resumo),
      topics: normalizeItems(parsed.topics || parsed.topicos, knownIds),
      decisions: normalizeItems(parsed.decisions || parsed.decisoes, knownIds),
      risks: normalizeItems(parsed.risks || parsed.riscos, knownIds),
      openQuestions: normalizeItems(
        parsed.openQuestions || parsed.perguntasAbertas,
        knownIds
      ),
      commitments: normalizeItems(
        parsed.commitments || parsed.compromissos,
        knownIds,
        { fields: ["what", "who"] }
      ),
      segmentCount,
      updatedAt: now
    };
  }

  function formatTranscript(transcript) {
    const lines = transcript.map((item) =>
      `[${item.id}] ${item.speaker || "Desconhecido"}: ${normalizeText(item.text)}`
    );

    let text = lines.join("\n");
    if (text.length > MAX_TRANSCRIPT_CHARS) {
      text = `…(início omitido)…\n${text.slice(-MAX_TRANSCRIPT_CHARS)}`;
    }

    return text;
  }

  const STATE_SYSTEM = [
    "Você acompanha uma reunião em andamento e mantém um estado estruturado dela.",
    "O transcript é a ÚNICA fonte factual: não invente fatos, nomes ou números.",
    "Toda inferência deve citar em `evidence` os ids das falas que a sustentam.",
    "Se não houver evidência para um item, não inclua o item.",
    "Responda SOMENTE com JSON válido, sem cercas de código e sem comentários.",
    "Formato:",
    '{"summary":"","topics":[{"text":"","evidence":["id"]}],',
    '"decisions":[{"text":"","evidence":["id"]}],',
    '"risks":[{"text":"","evidence":["id"]}],',
    '"openQuestions":[{"text":"","evidence":["id"]}],',
    '"commitments":[{"what":"","who":"","evidence":["id"]}]}',
    "Escreva em português do Brasil. Seja conciso e específico."
  ].join("\n");

  function buildStatePrompt(transcript, previousState = null) {
    const previous = previousState?.summary
      ? `Resumo anterior (atualize de forma incremental, não recomece):\n${previousState.summary}\n\n`
      : "";

    return {
      system: STATE_SYSTEM,
      messages: [{
        role: "user",
        content: `${previous}Transcript da reunião:\n${formatTranscript(transcript)}`
      }]
    };
  }

  const INSIGHT_SYSTEM = [
    "Você é um copiloto silencioso em uma reunião ao vivo.",
    "Só fale quando tiver algo REALMENTE útil e específico para o usuário contribuir.",
    "Silêncio é a resposta certa na maior parte do tempo.",
    "Não repita o que já foi dito, não resuma, não parabenize, não faça small talk.",
    `No máximo ${MAX_INSIGHTS_PER_BLOCK} itens por vez.`,
    "Responda SOMENTE com JSON válido, sem cercas de código.",
    'Formato: {"insights":[{"kind":"pergunta|ponto|risco","text":"","why":"","evidence":["id"]}]}',
    'Se nada for relevante, responda exatamente {"insights":[]}.',
    "Escreva em português do Brasil."
  ].join("\n");

  function buildInsightPrompt(recentSegments, state = null) {
    const context = state?.summary
      ? `Contexto da reunião até agora:\n${state.summary}\n\n`
      : "";

    return {
      system: INSIGHT_SYSTEM,
      messages: [{
        role: "user",
        content: `${context}Últimas falas:\n${formatTranscript(recentSegments)}`
      }]
    };
  }

  function parseInsights(raw, { transcript = [], seen = [] } = {}) {
    const parsed = extractJson(raw);
    if (!parsed) return [];

    const knownIds = new Set(
      transcript.map((item) => normalizeText(item?.id)).filter(Boolean)
    );
    const seenKeys = new Set(seen.map(dedupeKey));
    const list = Array.isArray(parsed) ? parsed : parsed.insights;
    if (!Array.isArray(list)) return [];

    const insights = [];

    for (const entry of list) {
      const text = normalizeText(entry?.text);
      if (!text) continue;

      const key = dedupeKey(text);
      if (seenKeys.has(key)) continue;
      seenKeys.add(key);

      insights.push({
        kind: ["pergunta", "ponto", "risco"].includes(entry?.kind)
          ? entry.kind
          : "ponto",
        text,
        why: normalizeText(entry?.why),
        evidence: normalizeEvidence(entry?.evidence, knownIds)
      });

      if (insights.length >= MAX_INSIGHTS_PER_BLOCK) break;
    }

    return insights;
  }

  // Analisa em blocos, não a cada fala: evita spam e chamada de modelo por linha.
  function shouldAnalyze({
    transcriptLength = 0,
    analyzedCount = 0,
    blockSize = 8
  } = {}) {
    if (transcriptLength <= 0) return false;
    if (transcriptLength < analyzedCount) return true;
    return transcriptLength - analyzedCount >= blockSize;
  }

  function buildChatPrompt(question, transcript, state = null) {
    const stateText = state?.summary
      ? `Estado atual da reunião:\n${state.summary}\n\n`
      : "";

    return {
      system: [
        "Você é o copiloto de reunião do usuário.",
        "Responda com base no transcript; se a resposta não estiver nele, diga isso.",
        "Seja direto e conciso. Português do Brasil."
      ].join(" "),
      messages: [{
        role: "user",
        content: `${stateText}Transcript:\n${formatTranscript(transcript)}\n\nPergunta: ${question}`
      }]
    };
  }

  const SHORTCUTS = [
    { id: "perguntar", label: "Perguntar", prompt: "Que pergunta forte eu poderia fazer agora nesta reunião?" },
    { id: "responder", label: "Responder", prompt: "Como eu poderia responder ao último ponto levantado?" },
    { id: "explicar", label: "Explicar", prompt: "Explique o último ponto técnico discutido, de forma simples." },
    { id: "contestar", label: "Contestar", prompt: "Qual é o contra-argumento mais forte ao que foi dito?" },
    { id: "riscos", label: "Riscos", prompt: "Quais riscos ou pontos cegos aparecem no que foi discutido?" },
    { id: "resumo", label: "Resumo", prompt: "Resuma a reunião até agora em tópicos objetivos." },
    { id: "minha-vez", label: "Minha vez", prompt: "É a minha vez de falar. O que vale eu dizer agora, em 2 ou 3 frases?" }
  ];

  return {
    EMPTY_STATE,
    MAX_INSIGHTS_PER_BLOCK,
    SHORTCUTS,
    buildChatPrompt,
    buildInsightPrompt,
    buildStatePrompt,
    emptyState,
    extractJson,
    formatTranscript,
    parseInsights,
    parseState,
    shouldAnalyze
  };
});

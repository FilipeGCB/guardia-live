const assert = require("node:assert/strict");
const test = require("node:test");
const {
  normalizeSettings,
  createNdjsonParser,
  createSseParser,
  extractOllamaToken,
  extractAnthropicToken,
  extractOpenAiToken,
  createClient,
  LlmError
} = require("../llm.js");

test("normalizeSettings aplica local-first por padrão", () => {
  const settings = normalizeSettings({});
  assert.equal(settings.provider, "ollama");
  assert.equal(settings.audioFallback, false);
  assert.equal(settings.retentionDays, 0);
});

test("normalizeSettings cai para ollama se provider externo não tem chave", () => {
  const settings = normalizeSettings({ provider: "anthropic", externalApiKey: "" });
  assert.equal(settings.provider, "ollama");
});

test("normalizeSettings mantém provider externo quando há chave", () => {
  const settings = normalizeSettings({ provider: "anthropic", externalApiKey: "sk-x" });
  assert.equal(settings.provider, "anthropic");
});

test("normalizeSettings limita proactiveBlockSize e retentionDays", () => {
  const settings = normalizeSettings({ proactiveBlockSize: 999, retentionDays: -5 });
  assert.equal(settings.proactiveBlockSize, 40);
  assert.equal(settings.retentionDays, 0);
});

test("createNdjsonParser reconstrói tokens do Ollama entre chunks", () => {
  const tokens = [];
  const parser = createNdjsonParser((obj) => tokens.push(extractOllamaToken(obj)));

  parser.push('{"message":{"content":"Ol');
  parser.push('á"}}\n{"message":{"content":" mundo"}}\n');
  parser.flush();

  assert.deepEqual(tokens, ["Olá", " mundo"]);
});

test("createNdjsonParser ignora linha incompleta final sem quebrar as anteriores", () => {
  const tokens = [];
  const parser = createNdjsonParser((obj) => tokens.push(extractOllamaToken(obj)));
  parser.push('{"message":{"content":"a"}}\n{"message":{"content":"b"');
  parser.flush();
  assert.deepEqual(tokens, ["a"]);
});

test("createSseParser extrai delta de texto da Anthropic e ignora [DONE]", () => {
  const tokens = [];
  const parser = createSseParser((obj) => {
    const token = extractAnthropicToken(obj);
    if (token) tokens.push(token);
  });

  parser.push('data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Oi"}}\n');
  parser.push("data: [DONE]\n");
  parser.flush();

  assert.deepEqual(tokens, ["Oi"]);
});

test("createSseParser extrai delta da OpenAI", () => {
  const tokens = [];
  const parser = createSseParser((obj) => {
    const token = extractOpenAiToken(obj);
    if (token) tokens.push(token);
  });

  parser.push('data: {"choices":[{"delta":{"content":"Oi"}}]}\n');
  parser.flush();

  assert.deepEqual(tokens, ["Oi"]);
});

test("createClient.chat rejeita quando nenhum modelo Ollama foi selecionado", async () => {
  const client = createClient({
    settings: { provider: "ollama", model: "" },
    fetch: async () => { throw new Error("não deveria chamar fetch"); }
  });

  await assert.rejects(
    () => client.chat({ messages: [{ role: "user", content: "oi" }] }),
    (error) => error instanceof LlmError && error.kind === "sem-modelo"
  );
});

test("createClient.chat streama tokens via onToken e resolve o texto completo", async () => {
  const encoder = new TextEncoder();
  const chunks = [
    '{"message":{"content":"Ol'.slice(0),
    'á"}}\n',
    '{"message":{"content":" mundo"},"done":true}\n'
  ];

  const fakeFetch = async () => ({
    ok: true,
    body: {
      getReader() {
        let index = 0;
        return {
          async read() {
            if (index >= chunks.length) return { done: true, value: undefined };
            const value = encoder.encode(chunks[index]);
            index += 1;
            return { done: false, value };
          }
        };
      }
    }
  });

  const client = createClient({
    settings: { provider: "ollama", model: "qwen3.8:27b" },
    fetch: fakeFetch
  });

  const seen = [];
  const full = await client.chat({
    messages: [{ role: "user", content: "oi" }],
    onToken: (token) => seen.push(token)
  });

  assert.deepEqual(seen, ["Olá", " mundo"]);
  assert.equal(full, "Olá mundo");
});

test("createClient.chat traduz falha de rede em LlmError offline com dica", async () => {
  const client = createClient({
    settings: { provider: "ollama", model: "qwen3.8:27b", ollamaUrl: "http://localhost:11434" },
    fetch: async () => { throw new TypeError("fetch failed"); }
  });

  await assert.rejects(
    () => client.chat({ messages: [{ role: "user", content: "oi" }] }),
    (error) => error instanceof LlmError && error.kind === "offline" && /OLLAMA_ORIGINS/.test(error.hint)
  );
});

test("createClient.listModels mapeia a resposta do /api/tags", async () => {
  const client = createClient({
    settings: { provider: "ollama" },
    fetch: async () => ({
      ok: true,
      json: async () => ({ models: [{ name: "qwen3.8:27b", size: 123 }] })
    })
  });

  const models = await client.listModels();
  assert.deepEqual(models, [{ name: "qwen3.8:27b", size: 123 }]);
});

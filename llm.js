(function (root, factory) {
  const api = factory();

  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.MeetingLlm = api;
  }
})(globalThis, function () {
  const DEFAULT_OLLAMA_URL = "http://localhost:11434";
  const DEFAULT_WHISPER_URL = "http://localhost:9000";

  // Local-first: provider local por padrão, nada externo, áudio desligado.
  const DEFAULT_SETTINGS = {
    provider: "ollama",
    ollamaUrl: DEFAULT_OLLAMA_URL,
    model: "",
    proactive: true,
    proactiveBlockSize: 8,
    meetingStateEnabled: true,
    audioFallback: false,
    whisperUrl: DEFAULT_WHISPER_URL,
    externalProvider: "",
    externalApiKey: "",
    externalModel: "",
    retentionDays: 0,
    debug: false
  };

  const PROVIDERS = ["ollama", "anthropic", "openai"];

  function normalizeSettings(stored) {
    const settings = { ...DEFAULT_SETTINGS, ...(stored || {}) };

    if (!PROVIDERS.includes(settings.provider)) {
      settings.provider = DEFAULT_SETTINGS.provider;
    }
    if (settings.externalProvider && !PROVIDERS.includes(settings.externalProvider)) {
      settings.externalProvider = "";
    }

    settings.ollamaUrl = trimUrl(settings.ollamaUrl) || DEFAULT_OLLAMA_URL;
    settings.whisperUrl = trimUrl(settings.whisperUrl) || DEFAULT_WHISPER_URL;
    settings.proactiveBlockSize = clampInteger(settings.proactiveBlockSize, 3, 40, 8);
    settings.retentionDays = clampInteger(settings.retentionDays, 0, 365, 0);

    for (const flag of ["proactive", "meetingStateEnabled", "audioFallback", "debug"]) {
      settings[flag] = settings[flag] === true;
    }

    // Um provider externo só vale com chave; sem chave, cai para o local.
    if (settings.provider !== "ollama" && !settings.externalApiKey) {
      settings.provider = "ollama";
    }

    return settings;
  }

  function trimUrl(value) {
    return String(value ?? "").trim().replace(/\/+$/, "");
  }

  function clampInteger(value, min, max, fallback) {
    const parsed = Number.parseInt(value, 10);
    if (Number.isNaN(parsed)) return fallback;
    return Math.min(Math.max(parsed, min), max);
  }

  // Parser incremental de NDJSON (formato de streaming do Ollama).
  function createNdjsonParser(onObject) {
    let buffer = "";

    function consume(line) {
      const trimmed = line.trim();
      if (!trimmed) return;

      try {
        onObject(JSON.parse(trimmed));
      } catch (error) {
        // Linha ainda incompleta ou ruído: ignora sem derrubar o stream.
      }
    }

    return {
      push(chunk) {
        buffer += chunk;
        let index = buffer.indexOf("\n");

        while (index >= 0) {
          consume(buffer.slice(0, index));
          buffer = buffer.slice(index + 1);
          index = buffer.indexOf("\n");
        }
      },
      flush() {
        const rest = buffer;
        buffer = "";
        consume(rest);
      }
    };
  }

  // Parser incremental de SSE (`data: {...}`), usado pelos providers externos.
  function createSseParser(onEvent) {
    let buffer = "";

    function consume(line) {
      const trimmed = line.trim();
      if (!trimmed.startsWith("data:")) return;

      const payload = trimmed.slice(5).trim();
      if (!payload || payload === "[DONE]") return;

      try {
        onEvent(JSON.parse(payload));
      } catch (error) {
        // Evento parcial: ignora.
      }
    }

    return {
      push(chunk) {
        buffer += chunk;
        let index = buffer.indexOf("\n");

        while (index >= 0) {
          consume(buffer.slice(0, index));
          buffer = buffer.slice(index + 1);
          index = buffer.indexOf("\n");
        }
      },
      flush() {
        const rest = buffer;
        buffer = "";
        consume(rest);
      }
    };
  }

  function extractOllamaToken(payload) {
    return payload?.message?.content || payload?.response || "";
  }

  function extractAnthropicToken(payload) {
    if (payload?.type !== "content_block_delta") return "";
    return payload?.delta?.text || "";
  }

  function extractOpenAiToken(payload) {
    return payload?.choices?.[0]?.delta?.content || "";
  }

  async function readStream(response, parser) {
    const body = response.body;

    if (!body?.getReader) {
      // Ambiente sem streaming (ou resposta já materializada): lê de uma vez.
      parser.push(await response.text());
      parser.flush();
      return;
    }

    const reader = body.getReader();
    const decoder = new TextDecoder();

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      parser.push(decoder.decode(value, { stream: true }));
    }

    parser.push(decoder.decode());
    parser.flush();
  }

  class LlmError extends Error {
    constructor(message, { kind = "erro", hint = "" } = {}) {
      super(message);
      this.name = "LlmError";
      this.kind = kind;
      this.hint = hint;
    }
  }

  const OLLAMA_OFFLINE_HINT =
    "Ollama não respondeu. Confirme que está rodando (`ollama serve`) e que " +
    "aceita a extensão: OLLAMA_ORIGINS=\"chrome-extension://*\" ollama serve";

  function extensionContext() {
    let runtimeId = "";
    let origin = "";

    try {
      runtimeId = String(
        globalThis.chrome?.runtime?.id || globalThis.browser?.runtime?.id || ""
      );
    } catch (error) {
      // Contexto da extensão pode já estar sendo invalidado.
    }

    try {
      origin = String(globalThis.location?.origin || "");
    } catch (error) {
      // Alguns contextos de worker não expõem location.
    }

    if (origin === "null") origin = "";
    if (!origin && runtimeId) origin = `chrome-extension://${runtimeId}`;

    return { runtimeId, origin };
  }

  function extensionContextHint() {
    const { runtimeId, origin } = extensionContext();
    const details = [
      runtimeId ? `chrome.runtime.id=${runtimeId}` : "chrome.runtime.id indisponível",
      origin ? `origin=${origin}` : "origin indisponível"
    ];
    return details.join(", ");
  }

  const OLLAMA_ORIGIN_HINT =
    "Ollama recusou a origem da extensão. No Linux, configure " +
    "OLLAMA_ORIGINS=\"chrome-extension://*\" de forma persistente e reinicie o Ollama.";

  function ollamaOriginHint() {
    return `${OLLAMA_ORIGIN_HINT} Diagnóstico: ${extensionContextHint()}.`;
  }

  function responseDetail(detail) {
    const text = String(detail ?? "").trim();
    return text ? ` ${text.slice(0, 200)}` : "";
  }

  async function describeOllamaResponse(response, operation) {
    const detail = await response.text().catch(() => "");
    const status = Number(response.status);
    const statusLabel = Number.isFinite(status) ? status : String(response.status);

    if (status === 403) {
      return new LlmError(
        `Ollama recusou a origem/permissão (HTTP 403) ${operation}.${responseDetail(detail)}`,
        { kind: "origem", hint: ollamaOriginHint() }
      );
    }

    return new LlmError(
      `Ollama respondeu HTTP ${statusLabel} ${operation}.${responseDetail(detail)}`,
      {
        kind: "http",
        hint: `O servidor Ollama respondeu, mas não aceitou a requisição. Diagnóstico: ${extensionContextHint()}.`
      }
    );
  }

  function describeFetchFailure(error, settings) {
    if (error?.name === "AbortError") {
      return new LlmError("Geração cancelada.", { kind: "cancelado" });
    }

    return new LlmError(
      `Não foi possível falar com o Ollama em ${settings.ollamaUrl}.`,
      { kind: "offline", hint: OLLAMA_OFFLINE_HINT }
    );
  }

  function createClient({ settings: rawSettings, fetch: fetchImpl } = {}) {
    const settings = normalizeSettings(rawSettings);
    const doFetch = fetchImpl || globalThis.fetch?.bind(globalThis);

    async function listModels() {
      if (!doFetch) return [];

      let response;
      try {
        response = await doFetch(`${settings.ollamaUrl}/api/tags`, {
          method: "GET"
        });
      } catch (error) {
        throw describeFetchFailure(error, settings);
      }

      if (!response.ok) {
        throw await describeOllamaResponse(response, "ao listar modelos");
      }

      const payload = await response.json();
      return (payload?.models || [])
        .map((entry) => ({
          name: entry?.name || entry?.model || "",
          size: entry?.size || 0
        }))
        .filter(({ name }) => name);
    }

    function requestFor(messages, { system, model, stream }) {
      const chosenModel = model || settings.model;

      if (settings.provider === "anthropic") {
        return {
          url: "https://api.anthropic.com/v1/messages",
          init: {
            method: "POST",
            headers: {
              "content-type": "application/json",
              "x-api-key": settings.externalApiKey,
              "anthropic-version": "2023-06-01",
              "anthropic-dangerous-direct-browser-access": "true"
            },
            // max_tokens limita thinking + texto; Opus 5 pensa por padrão,
            // então um teto baixo trunca a resposta no meio.
            body: JSON.stringify({
              model: settings.externalModel || "claude-opus-5",
              max_tokens: 16000,
              stream,
              ...(system ? { system } : {}),
              messages
            })
          },
          extract: extractAnthropicToken,
          parser: createSseParser
        };
      }

      if (settings.provider === "openai") {
        return {
          url: "https://api.openai.com/v1/chat/completions",
          init: {
            method: "POST",
            headers: {
              "content-type": "application/json",
              authorization: `Bearer ${settings.externalApiKey}`
            },
            body: JSON.stringify({
              model: settings.externalModel || "gpt-4o-mini",
              stream,
              messages: system
                ? [{ role: "system", content: system }, ...messages]
                : messages
            })
          },
          extract: extractOpenAiToken,
          parser: createSseParser
        };
      }

      return {
        url: `${settings.ollamaUrl}/api/chat`,
        init: {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({
            model: chosenModel,
            stream,
            messages: system
              ? [{ role: "system", content: system }, ...messages]
              : messages
          })
        },
        extract: extractOllamaToken,
        parser: createNdjsonParser
      };
    }

    async function chat({
      messages,
      system = "",
      model = "",
      signal = undefined,
      onToken = () => {}
    }) {
      if (!doFetch) throw new LlmError("fetch indisponível neste contexto.");
      if (settings.provider === "ollama" && !(model || settings.model)) {
        throw new LlmError("Nenhum modelo selecionado.", {
          kind: "sem-modelo",
          hint: "Escolha um modelo do Ollama na aba Config."
        });
      }

      const request = requestFor(messages, { system, model, stream: true });
      let response;

      try {
        response = await doFetch(request.url, { ...request.init, signal });
      } catch (error) {
        if (settings.provider !== "ollama") {
          if (error?.name === "AbortError") {
            throw new LlmError("Geração cancelada.", { kind: "cancelado" });
          }
          throw new LlmError(`Falha ao falar com ${settings.provider}.`);
        }
        throw describeFetchFailure(error, settings);
      }

      if (!response.ok) {
        if (settings.provider === "ollama") {
          throw await describeOllamaResponse(response, "ao gerar resposta");
        }

        const detail = await response.text().catch(() => "");
        throw new LlmError(
          `${settings.provider} respondeu ${response.status}. ${detail.slice(0, 200)}`,
          {}
        );
      }

      let full = "";
      const parser = request.parser((payload) => {
        const token = request.extract(payload);
        if (!token) return;
        full += token;
        onToken(token);
      });

      await readStream(response, parser);
      return full;
    }

    return { settings, listModels, chat };
  }

  return {
    DEFAULT_OLLAMA_URL,
    DEFAULT_SETTINGS,
    DEFAULT_WHISPER_URL,
    LlmError,
    OLLAMA_OFFLINE_HINT,
    PROVIDERS,
    createClient,
    createNdjsonParser,
    createSseParser,
    extractAnthropicToken,
    extractOllamaToken,
    extractOpenAiToken,
    normalizeSettings
  };
});

(function (root, factory) {
  const api = factory();

  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.MeetingTranscript = api;
  }
})(globalThis, function () {
  // Adaptador do Teams: TODO seletor do DOM do Teams vive nestas duas listas.
  // A primeira que casar vence e fica memoizada na instância da source.
  const CAPTION_SELECTORS = [
    '[data-tid="closed-caption-text"]',
    '[data-tid="closed-caption-message"]',
    '[data-tid="closed-captions-text"]',
    '[data-tid="caption-text"]',
    '[class*="closedCaptionText"]',
    '[class*="closed-caption-text"]'
  ];
  const AUTHOR_SELECTORS = [
    '[data-tid="author"]',
    '[data-tid="closed-caption-author"]',
    '[data-tid="message-author-name"]',
    '[class*="authorName"]'
  ];
  const CAPTION_SELECTOR = CAPTION_SELECTORS[0];
  const AUTHOR_SELECTOR = AUTHOR_SELECTORS[0];
  const LOG_PREFIX = "[MeetingCopilot]";
  const CAPTURE_LOG_PREFIX = "[MeetingCopilot][CAPTURE]";
  const CAPTURE_DIAGNOSTIC_PATTERN = /caption|subtitle|transcript|author|speaker/i;
  const MAX_DIAGNOSTIC_ELEMENTS = 20;
  const MAX_DIAGNOSTIC_TEXT_LENGTH = 160;
  const MAX_ZERO_DIAGNOSTICS = 5;
  const ZERO_DIAGNOSTIC_INTERVAL_MS = 30000;
  const UNKNOWN_SPEAKER = "Desconhecido";
  const DEFAULT_DEBOUNCE_MS = 1500;
  let generatedId = 0;

  function normalizeText(value) {
    return String(value ?? "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeSpeaker(value) {
    return normalizeText(value) || UNKNOWN_SPEAKER;
  }

  function timestamp(value, fallbackNow = Date.now()) {
    if (value !== undefined && value !== null) {
      const date = new Date(value);
      if (!Number.isNaN(date.getTime())) return date.toISOString();
    }

    return new Date(fallbackNow).toISOString();
  }

  function makeId(prefix = "caption") {
    if (typeof globalThis.crypto?.randomUUID === "function") {
      return `${prefix}-${globalThis.crypto.randomUUID()}`;
    }

    generatedId += 1;
    return `${prefix}-${Date.now()}-${generatedId}`;
  }

  function infoLog(event, details) {
    globalThis.console?.log?.(`${LOG_PREFIX} ${event}`, details);
  }

  function diagnosticLog(event, details) {
    if (globalThis.__MEETING_COPILOT_DIAGNOSTICS__ !== true) return;
    globalThis.console?.log?.(`${LOG_PREFIX} ${event}`, details);
  }

  function captureDiagnosticLog(event, details) {
    if (globalThis.__MEETING_COPILOT_DIAGNOSTICS__ !== true) return;
    globalThis.console?.log?.(`${CAPTURE_LOG_PREFIX} ${event}`, details);
  }

  function readAttribute(element, name) {
    return element?.getAttribute?.(name) ?? null;
  }

  function nodeContains(container, node) {
    if (!container || !node) return false;
    if (container === node) return true;

    try {
      if (container.contains?.(node)) return true;
    } catch (error) {
      // Fallback para os fakes e para nós de DOM que não expõem contains.
    }

    let current = node.parentElement;
    while (current) {
      if (current === container) return true;
      current = current.parentElement;
    }

    return false;
  }

  function truncateDiagnosticText(value) {
    const text = normalizeText(value);
    if (text.length <= MAX_DIAGNOSTIC_TEXT_LENGTH) return text;
    return `${text.slice(0, MAX_DIAGNOSTIC_TEXT_LENGTH)}…`;
  }

  function describeElement(element) {
    return {
      tagName: element?.tagName || null,
      dataTid: readAttribute(element, "data-tid"),
      ariaLabel: readAttribute(element, "aria-label"),
      role: readAttribute(element, "role"),
      text: truncateDiagnosticText(element?.innerText || element?.textContent || "")
    };
  }

  function describeMutationCandidates(records) {
    const elements = [];

    for (const record of records) {
      if (record?.type === "characterData") {
        const parent = record.target?.parentElement;
        if (parent) elements.push(parent);
        continue;
      }

      for (const node of record?.addedNodes || []) {
        const element = node?.nodeType === 1 ? node : node?.parentElement;
        if (element) elements.push(element);
      }
    }

    const uniqueCandidates = [...new Map(
      elements
        .map((element) => ({
          ...describeElement(element),
          parentTagName: element.parentElement?.tagName || null,
          parentDataTid: readAttribute(element.parentElement, "data-tid")
        }))
        .filter((candidate) => candidate.text)
        .map((candidate) => [JSON.stringify(candidate), candidate])
    ).values()];

    return {
      candidates: uniqueCandidates.slice(0, MAX_DIAGNOSTIC_ELEMENTS),
      total: uniqueCandidates.length,
      truncated: uniqueCandidates.length > MAX_DIAGNOSTIC_ELEMENTS
    };
  }

  function normalizeObservation(value, now = Date.now()) {
    if (!value) return null;

    const text = normalizeText(value.text);
    if (!text) return null;

    const id = normalizeText(value.id) || makeId();

    return {
      id,
      observedAt: timestamp(value.observedAt, now),
      speaker: normalizeSpeaker(value.speaker),
      text
    };
  }

  class SegmentAssembler {
    constructor() {
      this.active = null;
      this.finalizedIds = new Set();
    }

    observe(value) {
      const item = normalizeObservation(value);
      if (!item || this.finalizedIds.has(item.id)) return [];

      if (!this.active) {
        this.active = this.createActive(item);
        return [];
      }

      if (this.active.id === item.id) {
        if (
          item.speaker !== UNKNOWN_SPEAKER &&
          this.active.speaker !== UNKNOWN_SPEAKER &&
          item.speaker !== this.active.speaker
        ) {
          const completed = this.finalizeActive(item.observedAt);
          this.active = this.createActive(item);
          return [completed];
        }

        if (this.active.speaker === UNKNOWN_SPEAKER) {
          this.active.speaker = item.speaker;
        }

        if (item.text !== this.active.text) {
          this.active.text = item.text;
          this.active.lastObservedAt = item.observedAt;
        }

        return [];
      }

      const completed = this.finalizeActive(item.observedAt);
      this.active = this.createActive(item);
      return [completed];
    }

    finalize(id, finalizedAt = new Date().toISOString()) {
      if (!this.active || this.active.id !== id) return null;
      return this.finalizeActive(finalizedAt);
    }

    flush(finalizedAt = new Date().toISOString()) {
      if (!this.active) return null;
      return this.finalizeActive(finalizedAt);
    }

    createActive(item) {
      return {
        id: item.id,
        speaker: item.speaker,
        text: item.text,
        startedAt: item.observedAt,
        lastObservedAt: item.observedAt
      };
    }

    finalizeActive(finalizedAt) {
      const segment = {
        id: this.active.id,
        speaker: this.active.speaker,
        text: this.active.text,
        startedAt: this.active.startedAt,
        finalizedAt: timestamp(finalizedAt)
      };

      this.finalizedIds.add(this.active.id);
      this.active = null;
      return segment;
    }
  }

  class TeamsCaptionSource {
    constructor({
      root = globalThis.document,
      debounceMs = DEFAULT_DEBOUNCE_MS,
      now = () => Date.now(),
      setTimeout: schedule = globalThis.setTimeout,
      clearTimeout: cancel = globalThis.clearTimeout,
      onObservation = () => {},
      onFinalized = () => {},
      onCaptionsUnavailable = () => {},
      onCaptionsRecovered = () => {}
    } = {}) {
      this.root = root;
      this.debounceMs = debounceMs;
      this.now = now;
      this.schedule = schedule.bind(globalThis);
      this.cancel = cancel.bind(globalThis);
      this.onObservation = onObservation;
      this.onFinalized = onFinalized;
      this.onCaptionsUnavailable = onCaptionsUnavailable;
      this.onCaptionsRecovered = onCaptionsRecovered;
      this.states = new Map();
      this.elementStates = new WeakMap();
      this.elementSnapshots = new WeakMap();
      this.mutationObserver = null;
      this.scanSequence = 0;
      this.captionSelector = null;
      this.zeroDiagnosticCount = 0;
      this.lastZeroDiagnosticAt = null;
      this.lastZeroMutationDiagnostic = null;
      // Captions são sempre a fonte primária; isto só sinaliza ao content
      // script que o fallback de áudio (Whisper) pode valer a pena tentar.
      this.captionsUnavailable = false;
    }

    safeQueryAll(selector) {
      try {
        return [...this.root.querySelectorAll(selector)];
      } catch (error) {
        return [];
      }
    }

    // Não assume que o seletor atual do Teams continua válido: tenta a lista
    // inteira e memoiza o que funcionou.
    resolveCaptions() {
      if (this.captionSelector) {
        const matches = this.safeQueryAll(this.captionSelector);
        if (matches.length) return matches;
      }

      for (const selector of CAPTION_SELECTORS) {
        if (selector === this.captionSelector) continue;

        const matches = this.safeQueryAll(selector);
        if (!matches.length) continue;

        const previous = this.captionSelector;
        this.captionSelector = selector;
        infoLog("seletor de caption resolvido", {
          selector,
          anterior: previous,
          count: matches.length
        });
        return matches;
      }

      return [];
    }

    start() {
      if (!this.root?.querySelectorAll) return;

      if (this.mutationObserver) {
        captureDiagnosticLog("start ignorado", {
          reason: "observer já iniciado"
        });
        this.scan();
        return;
      }

      const MutationObserverClass = globalThis.MutationObserver;
      const observationRoot = this.root;

      if (MutationObserverClass && observationRoot) {
        this.mutationObserver = new MutationObserverClass((records = []) => {
          captureDiagnosticLog("mutation recebida", {
            records: records.length
          });
          this.scan({
            mutationDiagnostic: describeMutationCandidates(records),
            mutationRecords: records
          });
        });
        this.mutationObserver.observe(observationRoot, {
          subtree: true,
          childList: true,
          characterData: true
        });
        captureDiagnosticLog("observer iniciado", {
          root: observationRoot === this.root ? "root" : "outro",
          rootNodeName: observationRoot.nodeName || observationRoot.tagName || null
        });
      } else {
        captureDiagnosticLog("observer indisponível", {
          hasMutationObserver: Boolean(MutationObserverClass),
          hasObservationRoot: Boolean(observationRoot)
        });
      }

      // Tudo que já estava na página é apenas o baseline. A próxima mutation
      // que alterar uma caption cria/continua o lifecycle a partir do texto
      // observado, sem reemitir o histórico que o bootstrap encontrou.
      this.scan({ bootstrap: true });
    }

    stop() {
      this.mutationObserver?.disconnect();
      this.mutationObserver = null;

      for (const state of this.states.values()) {
        this.cancelStateTimer(state);
      }

      this.states.clear();
    }

    scan({
      mutationDiagnostic = null,
      mutationRecords = undefined,
      bootstrap = false
    } = {}) {
      const scanId = ++this.scanSequence;
      captureDiagnosticLog("scan iniciado", {
        scanId,
        selector: CAPTION_SELECTOR
      });

      if (!this.root?.querySelectorAll) {
        captureDiagnosticLog("scan descartado", {
          scanId,
          reason: "root sem querySelectorAll"
        });
        return;
      }

      const captions = this.resolveCaptions();
      captureDiagnosticLog("seletor principal", {
        scanId,
        selector: this.captionSelector || CAPTION_SELECTOR,
        count: captions.length
      });

      if (captions.length === 0) {
        this.inspectZeroCaptionResult(scanId);
        this.logZeroMutationDiagnostic(scanId, mutationDiagnostic);

        if (!this.captionsUnavailable && this.zeroDiagnosticCount >= MAX_ZERO_DIAGNOSTICS) {
          this.captionsUnavailable = true;
          infoLog("captions indisponíveis", { scanId, tentativas: this.zeroDiagnosticCount });
          this.onCaptionsUnavailable();
        }
      } else if (this.captionsUnavailable) {
        this.captionsUnavailable = false;
        infoLog("captions voltaram", { scanId, count: captions.length });
        this.onCaptionsRecovered();
      }

      const candidates = bootstrap
        ? captions
        : this.selectMutationCaptions(captions, mutationRecords);
      const seen = new Set(captions);
      const claimedStates = new Set();
      const currentTime = this.now();
      const observedAt = timestamp(currentTime, currentTime);

      for (const caption of candidates) {
        const text = readCaptionText(caption);
        const speaker = readSpeaker(caption);
        const candidateDetails = {
          scanId,
          tagName: caption?.tagName || null,
          dataTid: readAttribute(caption, "data-tid"),
          speaker,
          text
        };

        if (!text) {
          captureDiagnosticLog("candidato", {
            ...candidateDetails,
            discarded: true,
            reason: "texto vazio após normalização"
          });
          continue;
        }

        if (bootstrap) {
          this.elementSnapshots.set(caption, { text, speaker });
          captureDiagnosticLog("candidato", {
            ...candidateDetails,
            discarded: true,
            reason: "baseline do bootstrap",
            decision: "baseline"
          });
          continue;
        }

        let state = this.elementStates.get(caption);

        if (state && this.states.get(state.id) !== state) {
          state = null;
        }

        if (
          state &&
          isKnownSpeaker(speaker) &&
          isKnownSpeaker(state.speaker) &&
          speaker !== state.speaker
        ) {
          this.finalizeState(state, observedAt);
          state = null;
        }

        const previousSnapshot = this.elementSnapshots.get(caption);
        if (
          !state &&
          previousSnapshot?.text === text &&
          previousSnapshot?.speaker === speaker
        ) {
          captureDiagnosticLog("candidato", {
            ...candidateDetails,
            discarded: true,
            reason: "caption já observada sem alteração",
            decision: "duplicate"
          });
          continue;
        }

        if (!state) {
          state = this.findRecreatedState(text, speaker, seen, claimedStates);
        }

        if (!state) {
          const duplicateState = [...claimedStates].find(
            (candidate) =>
              candidate.text === text &&
              (!isKnownSpeaker(speaker) ||
                !isKnownSpeaker(candidate.speaker) ||
                candidate.speaker === speaker)
          );

          if (duplicateState) {
            captureDiagnosticLog("candidato", {
              ...candidateDetails,
              stateId: duplicateState.id,
              discarded: true,
              reason: "lifecycle já associado nesta varredura",
              decision: "duplicate"
            });
            continue;
          }
        }

        if (!state) {
          state = this.createState(caption, speaker, text, observedAt);
          claimedStates.add(state);
          captureDiagnosticLog("candidato", {
            ...candidateDetails,
            stateId: state.id,
            discarded: false,
            reason: null,
            decision: "new"
          });
          diagnosticLog("observation", {
            id: state.id,
            speaker,
            text,
            decision: "new",
            generation: state.generation,
            recreated: false
          });
          this.elementSnapshots.set(caption, {
            text: state.text,
            speaker: state.speaker
          });
          this.emitObservation(state, observedAt);
          continue;
        }

        const wasRecreated = state.element !== caption;
        claimedStates.add(state);
        const previousText = state.text;
        const previousSpeaker = state.speaker;

        this.elementStates.set(caption, state);
        state.element = caption;

        if (state.speaker === UNKNOWN_SPEAKER && isKnownSpeaker(speaker)) {
          state.speaker = speaker;
        }

        const meaningfulChange =
          text !== previousText || state.speaker !== previousSpeaker;

        diagnosticLog("observation", {
          id: state.id,
          speaker,
          text,
          decision: meaningfulChange ? "change" : "duplicate",
          generation: state.generation,
          recreated: wasRecreated
        });

        captureDiagnosticLog("candidato", {
          ...candidateDetails,
          stateId: state.id,
          discarded: !meaningfulChange,
          reason: meaningfulChange ? null : "sem alteração de texto ou speaker",
          decision: meaningfulChange ? "change" : "duplicate",
          recreated: wasRecreated
        });

        if (meaningfulChange) {
          state.text = text;
          if (isKnownSpeaker(speaker)) state.speaker = speaker;
          this.elementSnapshots.set(caption, {
            text: state.text,
            speaker: state.speaker
          });
          this.emitObservation(state, observedAt);
        } else {
          this.elementSnapshots.set(caption, {
            text: state.text,
            speaker: state.speaker
          });
        }
      }
    }

    selectMutationCaptions(captions, records) {
      if (records === undefined) return captions;
      if (!records.length) return [];

      // Mantém compatibilidade com chamadas de teste/integração que só
      // informam que o root mudou. O snapshot abaixo ainda impede reemissão.
      if (records.some((record) => !record?.type)) return captions;

      return captions.filter((caption) =>
        records.some((record) => this.captionAffectedByMutation(caption, record))
      );
    }

    captionAffectedByMutation(caption, record) {
      const target = record?.target;
      if (!target) return false;

      if (
        target === caption ||
        nodeContains(caption, target) ||
        (caption.parentElement &&
          (target === caption.parentElement ||
            nodeContains(caption.parentElement, target)))
      ) {
        return true;
      }

      if (record.type !== "childList" && record.type !== "characterData") {
        return false;
      }

      return [...(record.addedNodes || [])].some((node) =>
        nodeContains(node, caption) || nodeContains(caption, node)
      );
    }

    logZeroMutationDiagnostic(scanId, diagnostic) {
      if (globalThis.__MEETING_COPILOT_DIAGNOSTICS__ !== true) return;
      if (!diagnostic?.candidates.length) return;

      const signature = JSON.stringify(diagnostic.candidates);
      if (signature === this.lastZeroMutationDiagnostic) return;
      this.lastZeroMutationDiagnostic = signature;

      captureDiagnosticLog("diagnóstico mutation com seletor zero", {
        scanId,
        ...diagnostic
      });
    }

    // Roda quando NENHUM seletor casa. Repete algumas vezes espaçadas porque
    // o usuário costuma ligar as legendas depois do bootstrap — é este dump
    // que permite corrigir CAPTION_SELECTORS em uma única rodada.
    inspectZeroCaptionResult(scanId) {
      if (this.zeroDiagnosticCount >= MAX_ZERO_DIAGNOSTICS) return;

      const currentTime = this.now();
      if (
        this.lastZeroDiagnosticAt !== null &&
        currentTime - this.lastZeroDiagnosticAt < ZERO_DIAGNOSTIC_INTERVAL_MS
      ) {
        return;
      }

      this.lastZeroDiagnosticAt = currentTime;
      this.zeroDiagnosticCount += 1;

      try {
        const dataTidElements = [...this.root.querySelectorAll("[data-tid]")];
        const matchingDataTids = [...new Set(
          dataTidElements
            .map((element) => readAttribute(element, "data-tid"))
            .filter((value) => value && CAPTURE_DIAGNOSTIC_PATTERN.test(value))
        )].sort();
        const openShadowRoots = dataTidElements
          .map((element) => element.shadowRoot)
          .filter(Boolean);
        const shadowMatchingDataTids = [...new Set(
          openShadowRoots
            .flatMap((shadowRoot) => [...shadowRoot.querySelectorAll("[data-tid]")])
            .map((element) => readAttribute(element, "data-tid"))
            .filter((value) => value && CAPTURE_DIAGNOSTIC_PATTERN.test(value))
        )].sort();

        infoLog("nenhum seletor de caption casou", {
          scanId,
          tentativa: this.zeroDiagnosticCount,
          selectors: CAPTION_SELECTORS,
          matchingDataTids,
          shadowMatchingDataTids,
          openShadowRootCount: openShadowRoots.length
        });

        this.inspectZeroCaptionElements(scanId);
      } catch (error) {
        infoLog("diagnóstico seletor zero falhou", {
          scanId,
          error: String(error?.message ?? error)
        });
      }
    }

    // Varredura pesada (`*` + leitura de texto): só com diagnóstico ligado.
    inspectZeroCaptionElements(scanId) {
      if (globalThis.__MEETING_COPILOT_DIAGNOSTICS__ !== true) return;

      const relatedElements = [...this.root.querySelectorAll("*")]
        .map(describeElement)
        .filter((element) => CAPTURE_DIAGNOSTIC_PATTERN.test(
          [element.dataTid, element.ariaLabel, element.role, element.text]
            .filter(Boolean)
            .join(" ")
        ));

      captureDiagnosticLog("diagnóstico seletor zero", {
        scanId,
        selectors: CAPTION_SELECTORS,
        relatedElementCount: relatedElements.length,
        relatedElements: relatedElements.slice(0, MAX_DIAGNOSTIC_ELEMENTS),
        truncated: relatedElements.length > MAX_DIAGNOSTIC_ELEMENTS
      });
    }

    createState(element, speaker, text, observedAt) {
      const state = {
        id: makeId(),
        element,
        speaker,
        text,
        lastObservedAt: observedAt,
        timer: null,
        generation: 0
      };

      this.states.set(state.id, state);
      this.elementStates.set(element, state);
      return state;
    }

    findRecreatedState(text, speaker, seen, claimedStates = new Set()) {
      const candidates = [...this.states.values()]
        .filter((state) => {
          if (seen.has(state.element)) return false;
          if (claimedStates.has(state)) return false;
          if (
            isKnownSpeaker(speaker) &&
            isKnownSpeaker(state.speaker) &&
            speaker !== state.speaker
          ) {
            return false;
          }

          return isTextContinuation(state.text, text);
        })
        .map((state) => ({
          state,
          exact: state.text === text,
          strength: Math.min(state.text.length, text.length)
        }))
        .sort((left, right) => {
          if (left.exact !== right.exact) return left.exact ? -1 : 1;
          if (left.strength !== right.strength) {
            return right.strength - left.strength;
          }
          return right.state.lastObservedAt.localeCompare(
            left.state.lastObservedAt
          );
        });

      const best = candidates[0];
      const runnerUp = candidates[1];
      const ambiguous = runnerUp &&
        runnerUp.exact === best.exact &&
        runnerUp.strength === best.strength &&
        runnerUp.state.lastObservedAt === best.state.lastObservedAt;

      return ambiguous ? null : best?.state || null;
    }

    emitObservation(state, observedAt) {
      state.lastObservedAt = observedAt;
      captureDiagnosticLog("emitObservation", {
        id: state.id,
        observedAt,
        speaker: state.speaker,
        text: state.text
      });
      diagnosticLog("onObservation", {
        id: state.id,
        speaker: state.speaker,
        text: state.text,
        generation: state.generation
      });
      this.onObservation({
        id: state.id,
        observedAt,
        speaker: state.speaker,
        text: state.text
      });
      this.scheduleFinalization(state);
    }

    scheduleFinalization(state) {
      diagnosticLog("scheduleFinalization", {
        id: state.id,
        timer: state.timer,
        generation: state.generation,
        debounceMs: this.debounceMs
      });
      this.cancelStateTimer(state);
      const generation = state.generation;

      state.timer = this.schedule(() => {
        diagnosticLog("timer fire", {
          id: state.id,
          generation,
          currentGeneration: state.generation,
          stateIsCurrent: this.states.get(state.id) === state
        });

        if (
          this.states.get(state.id) !== state ||
          state.generation !== generation
        ) {
          return;
        }

        const currentTime = this.now();
        this.finalizeState(state, timestamp(currentTime, currentTime));
      }, this.debounceMs);

      diagnosticLog("timer scheduled", {
        id: state.id,
        timer: state.timer,
        generation
      });
    }

    cancelStateTimer(state) {
      const timer = state.timer;
      const previousGeneration = state.generation;
      if (timer !== null) this.cancel(timer);
      state.timer = null;
      state.generation += 1;
      diagnosticLog("timer cancel", {
        id: state.id,
        timer,
        previousGeneration,
        generation: state.generation
      });
    }

    finalizeState(state, finalizedAt) {
      const stateIsCurrent = this.states.get(state.id) === state;
      diagnosticLog("finalizeState", {
        id: state.id,
        finalizedAt,
        generation: state.generation,
        stateIsCurrent
      });

      if (!stateIsCurrent) return;

      this.cancelStateTimer(state);
      this.states.delete(state.id);
      diagnosticLog("onFinalized", {
        id: state.id,
        finalizedAt
      });
      this.onFinalized({ id: state.id, finalizedAt });
    }
  }

  function readCaptionText(caption) {
    return normalizeText(caption?.innerText || caption?.textContent || "");
  }

  function isTextContinuation(oldText, newText) {
    return (
      oldText === newText ||
      oldText.startsWith(newText) ||
      newText.startsWith(oldText)
    );
  }

  function readSpeaker(caption) {
    let node = caption;

    while (node) {
      for (const selector of AUTHOR_SELECTORS) {
        let author = null;

        try {
          author = node.querySelector?.(selector) ?? null;
        } catch (error) {
          continue;
        }

        const speaker = normalizeText(
          author?.innerText || author?.textContent || ""
        );
        if (speaker) return speaker;
      }

      node = node.parentElement;
    }

    return UNKNOWN_SPEAKER;
  }

  function isKnownSpeaker(speaker) {
    return speaker && speaker !== UNKNOWN_SPEAKER;
  }

  return {
    AUTHOR_SELECTOR,
    AUTHOR_SELECTORS,
    CAPTION_SELECTOR,
    CAPTION_SELECTORS,
    DEFAULT_DEBOUNCE_MS,
    SegmentAssembler,
    TeamsCaptionSource,
    normalizeObservation,
    normalizeText,
    readSpeaker
  };
});

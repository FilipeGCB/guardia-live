// Documento offscreen: único lugar do MV3 onde dá para segurar um
// MediaStream de áudio da aba vivo. Roda separado do service worker (que
// não tem getUserMedia) e do content script (que não tem tabCapture).
//
// Fallback de áudio: só entra quando as legendas falham E o usuário ligou
// a opção. Nunca substitui captions como fonte primária.

const LOG_PREFIX = "[MeetingCopilot][audio]";
const CHUNK_MS = 8000;
const MIN_CHUNK_BYTES = 2000; // silêncio puro gera blobs minúsculos; não vale mandar pro Whisper

let stream = null;
let recorder = null;
let whisperUrl = "";
let sessionMeetingId = null;
let stopping = false;

function log(event, details) {
  console.log(`${LOG_PREFIX} ${event}`, details ?? "");
}

async function transcribeChunk(blob) {
  if (blob.size < MIN_CHUNK_BYTES) return "";

  const form = new FormData();
  form.append("file", blob, "chunk.webm");
  form.append("model", "whisper-1");
  form.append("response_format", "json");

  const response = await fetch(`${whisperUrl}/v1/audio/transcriptions`, {
    method: "POST",
    body: form
  });

  if (!response.ok) {
    throw new Error(`Whisper respondeu ${response.status}`);
  }

  const payload = await response.json().catch(() => null);
  return String(payload?.text ?? "").trim();
}

function startRecorder() {
  recorder = new MediaRecorder(stream, { mimeType: "audio/webm;codecs=opus" });

  recorder.ondataavailable = async (event) => {
    if (stopping || !event.data || event.data.size === 0) return;

    try {
      const text = await transcribeChunk(event.data);
      if (!text) return;

      chrome.runtime.sendMessage({
        type: "meetingCopilot:whisperSegment",
        text,
        meetingId: sessionMeetingId,
        capturedAt: new Date().toISOString()
      }).catch(() => {});

      log("segmento transcrito", { chars: text.length });
    } catch (error) {
      // Uma falha de rede não pode matar o loop: só pula este pedaço.
      log("falha ao transcrever pedaço", String(error?.message ?? error));
    }
  };

  recorder.onerror = (event) => {
    log("erro do MediaRecorder", String(event?.error?.message ?? event));
  };

  recorder.start(CHUNK_MS);
  log("gravação iniciada", { chunkMs: CHUNK_MS });
}

async function startCapture({ streamId, whisperUrl: url, meetingId }) {
  if (stream) return; // já capturando — idempotente

  stopping = false;
  whisperUrl = url;
  sessionMeetingId = meetingId ?? null;

  stream = await navigator.mediaDevices.getUserMedia({
    audio: {
      mandatory: {
        chromeMediaSource: "tab",
        chromeMediaSourceId: streamId
      }
    }
  });

  // Roteia o áudio de volta para a saída: a extensão não pode silenciar a
  // reunião do usuário só porque passou a escutar também.
  const playback = document.getElementById("playback");
  playback.srcObject = stream;
  playback.play().catch((error) => log("falha ao tocar áudio de volta", String(error)));

  startRecorder();
}

function stopCapture() {
  stopping = true;

  try {
    recorder?.stop();
  } catch (error) {
    // já parado
  }
  recorder = null;

  for (const track of stream?.getTracks?.() || []) track.stop();
  stream = null;

  const playback = document.getElementById("playback");
  if (playback) playback.srcObject = null;

  log("captura de áudio parada");
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "meetingCopilot:offscreen:start") {
    startCapture(message)
      .then(() => sendResponse({ ok: true }))
      .catch((error) => {
        log("falha ao iniciar captura", String(error?.message ?? error));
        sendResponse({ ok: false, error: String(error?.message ?? error) });
      });
    return true;
  }

  if (message?.type === "meetingCopilot:offscreen:stop") {
    stopCapture();
    sendResponse({ ok: true });
    return false;
  }

  return undefined;
});

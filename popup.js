const ext = globalThis.browser ?? globalThis.chrome;

const transcriptEl = document.getElementById("transcript");
const summaryEl = document.getElementById("summary");
const clearButton = document.getElementById("clear");

function formatTime(value) {
  try {
    return new Date(value).toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  } catch {
    return "";
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function reportStorageError(operation, error) {
  console.error(`❌ Falha ${operation}`, error);
}

async function clearTranscript() {
  const response = await ext.runtime.sendMessage({
    type: "meetingCopilot:clearTranscript",
    nextMeetingId: crypto.randomUUID()
  });
  if (!response?.ok) {
    throw new Error(response?.error || "Falha desconhecida no storage");
  }
}

async function render() {
  const result = await ext.storage.local.get("meetingTranscript");

  const transcript = result.meetingTranscript || [];

  summaryEl.textContent =
    `${transcript.length} fala${transcript.length === 1 ? "" : "s"} salva${transcript.length === 1 ? "" : "s"}`;

  if (transcript.length === 0) {
    transcriptEl.innerHTML =
      '<div id="empty">Nenhuma fala capturada ainda.</div>';
    return;
  }

  transcriptEl.innerHTML = transcript
    .map(item => `
      <div class="item">
        <div class="meta">
          ${escapeHtml(formatTime(item.time))}
          ·
          <span class="speaker">${escapeHtml(item.speaker)}</span>
        </div>

        <div class="text">
          ${escapeHtml(item.text)}
        </div>
      </div>
    `)
    .join("");

  transcriptEl.scrollTop = transcriptEl.scrollHeight;
}

clearButton.addEventListener("click", async () => {
  const confirmed = confirm("Apagar todo o histórico capturado?");

  if (!confirmed) return;

  try {
    await clearTranscript();

    await render();
  } catch (error) {
    reportStorageError("ao limpar transcript", error);
  }
});

ext.storage.onChanged.addListener((changes, areaName) => {
  if (
    areaName === "local" &&
    changes.meetingTranscript
  ) {
    render().catch((error) => {
      reportStorageError("ao atualizar popup", error);
    });
  }
});

render().catch((error) => {
  reportStorageError("ao carregar popup", error);
});

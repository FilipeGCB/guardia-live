const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const popupSource = fs.readFileSync(
  path.join(__dirname, "..", "popup.js"),
  "utf8"
);

async function settle() {
  for (let index = 0; index < 5; index += 1) await Promise.resolve();
  await new Promise((resolve) => setImmediate(resolve));
}

class FakeElement {
  constructor() {
    this.textContent = "";
    this.innerHTML = "";
    this.listeners = new Map();
  }

  addEventListener(type, listener) {
    this.listeners.set(type, listener);
  }

  async click() {
    await this.listeners.get("click")?.();
  }
}

async function loadPopup({ getError = null, setError = null } = {}) {
  const errors = [];
  const elements = {
    transcript: new FakeElement(),
    summary: new FakeElement(),
    clear: new FakeElement()
  };
  const storage = {
    local: {
      async get() {
        if (getError) throw getError;
        return { meetingTranscript: [] };
      },
      async set() {
        if (setError) throw setError;
      }
    },
    onChanged: { addListener() {} }
  };
  const runtime = {
    async sendMessage() {
      if (setError) throw setError;
      return { ok: true };
    }
  };
  const context = vm.createContext({
    browser: undefined,
    chrome: { storage, runtime },
    confirm: () => true,
    crypto: { randomUUID: () => "session-after-clear" },
    console: {
      error(...args) {
        errors.push(args);
      }
    },
    document: {
      getElementById(id) {
        return elements[id];
      }
    }
  });

  vm.runInContext(popupSource, context, { filename: "popup.js" });
  await settle();
  return { elements, errors };
}

test("Popup reporta falha de render sem unhandled rejection", async () => {
  const popup = await loadPopup({
    getError: new Error("Temporary popup read failure")
  });

  assert.equal(popup.errors.length, 1);
  assert.match(String(popup.errors[0][1]), /Temporary popup read failure/);
});

test("Popup reporta falha ao limpar sem rejeitar o handler", async () => {
  const popup = await loadPopup({
    setError: new Error("Temporary popup write failure")
  });

  await popup.elements.clear.click();
  await settle();

  assert.equal(popup.errors.length, 1);
  assert.match(String(popup.errors[0][1]), /Temporary popup write failure/);
});

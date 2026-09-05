# Guard.IA Live

**A local-first AI copilot for following meetings in real time, organizing context, and supporting the conversation as it happens.**

**PT-BR:** Copiloto local-first para acompanhar reuniões em tempo real, organizar contexto e apoiar a conversa enquanto ela acontece.

[Versão em português](README.md)

## In 10 seconds

Guard.IA Live is a browser extension that follows **visible meeting captions**, builds a local transcript, keeps session context, and provides a copilot in the side panel.

The flow that is actually implemented today is focused on **Microsoft Teams Web**. The project does not claim to be a universal capture system for every messenger.

## Why it exists

In long meetings, the problem is not only transcription. It is **not losing context** while the conversation evolves.

The project explores a copilot that follows what was said, keeps local history, and lets the user query the session without making a cloud service mandatory by default.

## What is implemented

- Teams caption capture with deduplication, lifecycle handling, and reinjection after extension reload;
- current transcript, history, sessions, Meeting State, and insights stored in `chrome.storage.local`;
- streaming chat and shortcuts in the side panel;
- **local Ollama as the default provider**;
- Anthropic/OpenAI as optional providers after explicit configuration;
- optional Whisper audio fallback on Chrome when captions fail;
- initial Firefox adaptation using `sidebar_action`.

Part of the technical UI naming is still **Meeting Copilot**. The public project name is Guard.IA Live without pretending that every legacy naming surface has already been migrated.

## Current state

**Functional prototype under active evolution.** The current baseline covers Teams Web and local storage, with important limitations:

- capture depends on Teams captions and current DOM selectors;
- a real Teams meeting and the complete live flow still require manual validation;
- audio fallback is not available on Firefox;
- the extension does not bundle a Whisper server build or AI models;
- generic capture for other messengers is not implemented yet.

## Privacy and local data

- captions, transcript, history, and settings remain in browser-local storage;
- the default provider does not send data to the internet;
- external providers are used only after explicit configuration and permission;
- audio fallback is disabled by default;
- **Delete everything now** removes extension data from the browser.

The repository does not contain real transcripts, private conversations, storage dumps, logs, credentials, valid tokens, or keys. Fixtures and placeholders are synthetic.

## How to test

### Requirements

- Node.js for tests;
- local Ollama at `http://localhost:11434` for the copilot.

```bash
OLLAMA_ORIGINS="chrome-extension://*" ollama serve
```

The optional Whisper server uses:

```text
POST /v1/audio/transcriptions
```

### Chrome

1. Open `chrome://extensions` and enable Developer mode.
2. Choose **Load unpacked** and select this folder.
3. Open Teams Web, enable captions, and open the extension panel.

### Firefox

1. Open `about:debugging#/runtime/this-firefox`.
2. Choose **Load Temporary Add-on** and select `manifest.json`.
3. Open the sidebar and select **Meeting Copilot**.

The current manifest contains Chrome-specific permissions (`sidePanel`, `offscreen`, `tabCapture`, and `side_panel`), so Firefox may display warnings.

## Automated tests

```bash
node --test
for file in ./*.js; do node --check "$file" || exit 1; done
```

The harnesses in `test/*.html` exercise capture and lifecycle against a real browser DOM when served locally.

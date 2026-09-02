# Guard.IA Live

Copiloto de IA em tempo real para reuniões, conversas e mensageiros.

## Estado atual

Este repositório contém uma extensão WebExtension MV3. O fluxo implementado
hoje captura legendas visíveis do Microsoft Teams Web, monta um transcript,
mantém histórico local e oferece um copiloto no painel lateral.

O nome técnico da extensão e parte da interface ainda é **Meeting Copilot**.
Esta publicação nomeia o projeto como Guard.IA Live sem alterar a
funcionalidade existente. A captura genérica de outros mensageiros ainda não
está implementada.

## O que está implementado

- Captura de captions do Teams com deduplicação, lifecycle e reinjeção após
  recarga da extensão.
- Transcript atual, histórico, sessões, Meeting State e insights persistidos
  em `chrome.storage.local`.
- Chat com streaming e atalhos no painel lateral.
- Ollama local como provider padrão; Anthropic/OpenAI são opt-in.
- Fallback opcional de áudio via Whisper no Chrome quando as captions falham.
- Adaptação inicial para Firefox com `sidebar_action`; o fallback de áudio
  não está disponível no Firefox.

## Requisitos

- Node.js para executar os testes.
- Ollama local em `http://localhost:11434` para o copiloto:

```bash
OLLAMA_ORIGINS="chrome-extension://*" ollama serve
```

Esse comando vale apenas para o processo iniciado no terminal. Para manter a
origem autorizada em instalações Linux gerenciadas pelo `systemd`:

```bash
sudo systemctl edit ollama
```

Adicione ao override:

```ini
[Service]
Environment="OLLAMA_ORIGINS=chrome-extension://*"
```

Depois recarregue e reinicie o serviço:

```bash
sudo systemctl daemon-reload
sudo systemctl restart ollama
systemctl show ollama --property=Environment
```

O valor `chrome-extension://*` autoriza extensões Chrome, incluindo a origem
da extensão exibida no diagnóstico de erro 403.

O servidor Whisper é opcional e usa o contrato
`POST /v1/audio/transcriptions`.

## Carregar para testar

### Chrome

1. Abra `chrome://extensions` e ative o modo do desenvolvedor.
2. Selecione **Carregar sem compactação** e escolha esta pasta.
3. Abra o Teams Web, habilite as captions e abra o painel da extensão.

### Firefox

1. Abra `about:debugging#/runtime/this-firefox`.
2. Selecione **Carregar extensão temporária** e escolha `manifest.json`.
3. Abra a sidebar pelo menu de Sidebars e selecione **Meeting Copilot**.

O manifesto atual contém chaves/permissões específicas do Chrome
(`sidePanel`, `offscreen`, `tabCapture` e `side_panel`); o Firefox pode
exibir avisos ao carregá-lo. Isso é uma limitação conhecida deste baseline.

## Privacidade e dados locais

- Captions, transcript, histórico e configurações ficam no armazenamento local
  do navegador.
- O provider padrão não envia dados para a internet.
- Providers externos só são usados após configuração explícita e solicitação
  da permissão correspondente.
- O fallback de áudio fica desligado por padrão.
- O botão **Apagar tudo agora** remove os dados da extensão no navegador.

Não há transcripts, conversas reais, dumps de armazenamento, logs, credenciais,
tokens válidos ou chaves neste repositório. Há apenas fixtures sintéticas e
placeholders de teste. Artefatos locais desse tipo são ignorados por
`.gitignore`.

## Testes

```bash
node --test
for file in ./*.js; do node --check "$file" || exit 1; done
```

Os harnesses em `test/*.html` exercitam a captura e o lifecycle contra DOM
real do navegador quando servidos localmente.

## Limitações conhecidas

- A captura depende das captions e dos seletores atuais do DOM do Teams.
- Uma reunião real do Teams e o fluxo completo ao vivo ainda exigem validação
  manual.
- A extensão não inclui build, servidor Whisper ou modelos de IA.

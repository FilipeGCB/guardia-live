# Guard.IA Live

**Copiloto local-first para acompanhar reuniões em tempo real, organizar contexto e apoiar a conversa enquanto ela acontece.**

**EN:** A local-first AI copilot for following meetings in real time, organizing context, and supporting the conversation as it happens.

[English version](README.en.md)

## Em 10 segundos

Guard.IA Live é uma extensão de navegador que acompanha **captions visíveis de reuniões**, monta um transcript local, mantém contexto da sessão e oferece um copiloto no painel lateral.

Hoje, o fluxo realmente implementado está concentrado no **Microsoft Teams Web**. O projeto não se apresenta como um capturador universal de qualquer mensageiro.

## Por que existe

Em reuniões longas, o problema não é só transcrever: é **não perder contexto** enquanto a conversa evolui.

O projeto explora um copiloto que acompanha o que foi dito, mantém histórico local e permite consultar a sessão sem tornar um serviço de nuvem obrigatório por padrão.

## O que está implementado

- captura de captions do Teams com deduplicação, lifecycle e reinjeção após recarga da extensão;
- transcript atual, histórico, sessões, Meeting State e insights persistidos em `chrome.storage.local`;
- chat com streaming e atalhos no painel lateral;
- **Ollama local como provider padrão**;
- Anthropic/OpenAI opcionais e somente após configuração explícita;
- fallback opcional de áudio via Whisper no Chrome quando captions falham;
- adaptação inicial para Firefox com `sidebar_action`.

O nome técnico de parte da interface ainda é **Meeting Copilot**. A publicação usa Guard.IA Live como nome do projeto sem fingir que todo o legado de naming já foi migrado.

## Estado atual

**Protótipo funcional em evolução.** O baseline atual cobre Teams Web e armazenamento local, mas ainda há limites importantes:

- a captura depende das captions e dos seletores atuais do DOM do Teams;
- uma reunião real do Teams e o fluxo completo ao vivo ainda exigem validação manual;
- o fallback de áudio não está disponível no Firefox;
- a extensão não inclui build de servidor Whisper nem modelos de IA;
- captura genérica de outros mensageiros ainda não está implementada.

## Privacidade e dados locais

- captions, transcript, histórico e configurações ficam no armazenamento local do navegador;
- o provider padrão não envia dados para a internet;
- providers externos só são usados após configuração explícita e solicitação da permissão correspondente;
- o fallback de áudio fica desligado por padrão;
- o botão **Apagar tudo agora** remove os dados da extensão no navegador.

O repositório não contém transcripts reais, conversas privadas, dumps de armazenamento, logs, credenciais, tokens válidos ou chaves. Fixtures e placeholders são sintéticos.

## Como testar

### Requisitos

- Node.js para executar os testes;
- Ollama local em `http://localhost:11434` para o copiloto.

```bash
OLLAMA_ORIGINS="chrome-extension://*" ollama serve
```

O servidor Whisper é opcional e usa o contrato:

```text
POST /v1/audio/transcriptions
```

### Chrome

1. Abra `chrome://extensions` e ative o modo do desenvolvedor.
2. Selecione **Carregar sem compactação** e escolha esta pasta.
3. Abra o Teams Web, habilite captions e abra o painel da extensão.

### Firefox

1. Abra `about:debugging#/runtime/this-firefox`.
2. Selecione **Carregar extensão temporária** e escolha `manifest.json`.
3. Abra a sidebar e selecione **Meeting Copilot**.

O manifesto atual contém permissões específicas do Chrome (`sidePanel`, `offscreen`, `tabCapture` e `side_panel`), então o Firefox pode exibir avisos.

## Testes automatizados

```bash
node --test
for file in ./*.js; do node --check "$file" || exit 1; done
```

Os harnesses em `test/*.html` exercitam captura e lifecycle contra DOM real do navegador quando servidos localmente.

# OpenWordCode

<p align="center">
  <img src="apps/word-addin/public/openwordcode-logo.png" alt="OpenWordCode logo" width="220" />
</p>

<p align="center"><strong>AI assistance for Microsoft Word</strong></p>

OpenWordCode puts an AI agent in Word. Ask it to read, explain, rewrite, or
edit your document while you stay in Word.

Created and maintained by **Liko** · [MIT License](LICENSE)

## Features

- Chat with the current document, selection, tables, and images.
- Rewrite, summarize, review, format, draft, and create document content.
- Review proposed edits before they reach Word, then apply or revert them.
- Use manual approval, automatic approval, or skip approvals.
- Attach or drag in PDFs and images, and paste images with `Ctrl+V`.
- Preview uploaded images and open them for a closer look.
- Connect multiple providers and switch providers and models from the chat.
- Sign in with supported OAuth flows or official local CLI sessions.
- Search the web, load AI skills, keep chat history, and start new tasks.
- See context usage and automatically compact long conversations.

OpenWordCode is in early development. Provider availability and Word features
depend on your account and installed Microsoft Word version.

## Requirements

- Windows 10 or Windows 11
- Microsoft Word Desktop with Office Add-ins enabled
- Node.js 20 or newer

## Install and run

```powershell
git clone https://github.com/LikoKiko/OpenWordCode.git
cd OpenWordCode
npm ci
Copy-Item .env.example .env
npm run cert:install
npm run dev
```

Keep that terminal running. The local services use these addresses:

| Service | Address |
| --- | --- |
| Word task pane | `https://localhost:3000` |
| Core | `http://127.0.0.1:10200` |
| Bridge | `http://127.0.0.1:10101` |

Open `https://localhost:3000` once in your browser and accept the local
development certificate warning. Then open a second terminal and run:

```powershell
npm run sideload
```

Open Word and choose **Home → Add-ins → OpenWordCode**.

When you are finished:

```powershell
npm run unsideload
```

## Use it in Word

1. Open **Settings** in the task pane and connect a provider.
2. Return to **Chat**, choose the provider and model you want to use.
3. Type a request, select text in Word, or attach a file.
4. Review the proposed change and apply it when it looks right.

Try prompts like:

```text
Summarize this document in five bullet points.
```

```text
Rewrite the selected paragraph to sound clearer and more professional.
```

```text
Redesign the selected table without changing its content.
```

API-key entry is intentionally not shown in the public Word interface. Use the
OpenWordCode Bridge, supported OAuth, or an official provider CLI session. See
[`docs/AUTHENTICATION.md`](docs/AUTHENTICATION.md) for provider setup details.

## Useful commands

| Command | What it does |
| --- | --- |
| `npm run dev` | Start the Core and Word task pane |
| `npm run doctor` | Check the local setup |
| `npm run providers` | List configured providers |
| `npm run models` | List available models |
| `npm test` | Run tests |
| `npm run typecheck` | Check TypeScript |
| `npm run build` | Build the task pane |

## Project layout

```text
apps/core/          Local Core, authentication, Bridge, and agent server
apps/word-addin/    React task pane and Office manifest
packages/agent-core Agent prompts and tools
packages/app-word   Word and in-memory document adapters
packages/providers  Provider connections and model discovery
packages/shared     Shared types and protocols
tests/              Automated tests
docs/               Authentication, privacy, and project notes
```

## Security and privacy

- `.env` and local credential data are ignored by Git. Never commit secrets.
- Credentials stay in the local Core credential store where possible.
- Cloud providers receive the document context and attachments needed for a
  request. Do not send sensitive documents unless you trust that provider.
- The local console tool is bounded and read-only in this development version.
- There is no hosted OpenWordCode backend or billing system in this repository.

Read [`SECURITY.md`](SECURITY.md) and [`docs/PRIVACY.md`](docs/PRIVACY.md)
before sharing the project with other users.

## Contributing

```powershell
npm ci
npm run typecheck
npm test
npm run build
```

Please read [`CONTRIBUTING.md`](CONTRIBUTING.md), open an issue for bugs or
ideas, and send a pull request for changes.

## Credits and license

OpenWordCode was created by **Liko** and is released under the [MIT License](LICENSE).

See [`ATTRIBUTIONS.md`](ATTRIBUTIONS.md) for third-party notices. Microsoft
Word, Office.js, and provider names belong to their respective owners.

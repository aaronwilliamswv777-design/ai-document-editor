# Doc Editing Application (MVP)

Dockerized web app for human-in-the-loop AI editing of Word documents (`.docx`):

- Upload a source `.docx` (immutable source copy for safety).
- Upload optional context files (`.docx`, `.txt`, `.md`).
- Prompt AI for targeted edits.
- Review clear highlighted diffs per proposed edit.
- Accept or reject each proposal individually.
- Download the current working copy as a new `.docx`.
- Optionally promote working copy to session source baseline with explicit confirmation.

## Stack

- Frontend: React + TypeScript + Vite
- Backend: Express + TypeScript
- Document parsing: `mammoth`
- Document export: `docx`
- Diff rendering: `diff`
- AI providers:
  - Anthropic (Claude)
  - Gemini
  - OpenRouter
  - Mock fallback (no API key needed)

## Project Structure

```text
.
├─ docker-compose.yml
├─ backend/
│  ├─ src/index.ts
│  ├─ src/services/aiService.ts
│  ├─ src/services/documentService.ts
│  └─ src/sessionStore.ts
└─ frontend/
   ├─ src/App.tsx
   ├─ src/api.ts
   └─ src/styles.css
```

## Run

1. Copy `.env.example` to `.env` and set provider API keys if needed.
2. Start:

```bash
docker compose up --build
```

3. Open `http://localhost:5173`.

## Notes and Current Constraints

- Source doc support is `.docx` only.
- Current parser/extractor focuses on raw paragraph text for MVP.
- Complex Word structures (tables, headers/footers, footnotes, tracked changes) are not preserved in this MVP flow.
- Session state is in-memory (single instance, non-persistent).

## API Endpoints (MVP)

- `POST /api/session`
- `DELETE /api/session/:id`
- `POST /api/session/:id/upload-source` (`multipart/form-data`, field `file`, `.docx`)
- `POST /api/session/:id/upload-context` (`multipart/form-data`, field `file`)
- `GET /api/session/:id/state`
- `POST /api/session/:id/propose-edits`
- `POST /api/session/:id/edits/:editId/decision`
- `POST /api/session/:id/promote-working`
- `GET /api/session/:id/download?variant=working|source`


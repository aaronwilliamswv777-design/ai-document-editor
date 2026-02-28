# Doc Editing Application (MVP)

Dockerized web app for human-in-the-loop AI editing of Word documents (`.docx`):

- Upload a source `.docx` (immutable source copy for safety).
- Upload optional context files (`.docx`, `.txt`, `.md`).
- Prompt AI for targeted edits.
- Switch to Grammar & Punctuation mode and run document analysis.
- Optional custom instructions in Grammar & Punctuation mode.
- Add provider API keys in a dedicated API Key menu in the UI.
- Review clear highlighted diffs per proposed edit.
- See grammar/punctuation detections highlighted in red in the formatted preview with hover explanations.
- Accept or reject each proposal individually.
- Accept all pending proposals in one action.
- Download the current working copy as a new `.docx`.
- Optionally promote working copy to session source baseline with explicit confirmation.

## Screenshots

### Main Editor View

![Main Editor View](Screenshots/Screenshot%202026-02-24%20191613.png)

### Preview and Diff View

![Preview and Diff View](Screenshots/Screenshot%202026-02-24%20191657.png)

## Stack

- Frontend: React + TypeScript + Vite
- Backend: Express + TypeScript
- Document parsing: `mammoth`
- Document mutation/preservation: `jszip` + OOXML XML patching
- Diff rendering: `diff`
- Formatted preview: `docx-preview`
- AI providers:
  - Anthropic (Claude)
  - Gemini
  - OpenRouter

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

1. Optional: copy `.env.example` to `.env` if you want to override default model IDs.
2. Start:

```bash
docker compose up --build
```

3. Open `http://localhost:5173`.
4. Enter your own provider API key in the UI (AI Settings) before loading models or running analysis.

## Notes and Current Constraints

- Source doc support is `.docx` only.
- Exports are generated from the original DOCX package with in-place text edits, preserving document-level structure and styling metadata.
- Formatted preview is rendered from the real working DOCX in-browser.
- Inline run-level styling inside heavily edited paragraphs may shift if an edit materially changes run boundaries.
- Session state is in-memory (single instance, non-persistent).

## API Endpoints (MVP)

- `POST /api/session`
- `DELETE /api/session/:id`
- `POST /api/session/:id/upload-source` (`multipart/form-data`, field `file`, `.docx`)
- `POST /api/session/:id/upload-context` (`multipart/form-data`, field `file`)
- `GET /api/session/:id/state`
- `POST /api/session/:id/propose-edits`
- `POST /api/session/:id/analyze-grammar`
- `POST /api/session/:id/edits/:editId/decision`
- `POST /api/session/:id/edits/accept-all`
- `POST /api/session/:id/promote-working`
- `GET /api/session/:id/download?variant=working|source`
- `GET /api/session/:id/preview?variant=working|source`

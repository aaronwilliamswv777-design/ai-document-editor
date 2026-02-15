import { ProposalBatch, SessionState } from "./types";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:8080";

async function parseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    let detail = "Request failed";
    try {
      const body = (await response.json()) as { error?: string };
      if (body.error) {
        detail = body.error;
      }
    } catch {
      // Keep default message.
    }
    throw new Error(detail);
  }
  return (await response.json()) as T;
}

export async function createSession(): Promise<{ id: string; createdAt: string }> {
  const response = await fetch(`${API_BASE}/api/session`, {
    method: "POST"
  });
  return parseJson(response);
}

export async function fetchState(sessionId: string): Promise<SessionState> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/state`);
  return parseJson(response);
}

export async function uploadSource(sessionId: string, file: File): Promise<void> {
  const form = new FormData();
  form.append("file", file);
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/upload-source`, {
    method: "POST",
    body: form
  });
  await parseJson(response);
}

export async function uploadContext(sessionId: string, file: File): Promise<void> {
  const form = new FormData();
  form.append("file", file);
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/upload-context`, {
    method: "POST",
    body: form
  });
  await parseJson(response);
}

export async function proposeEdits(
  sessionId: string,
  payload: { prompt: string; provider: "anthropic" | "gemini" | "openrouter" | "mock"; model?: string }
): Promise<ProposalBatch> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/propose-edits`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  return parseJson(response);
}

export async function decideEdit(
  sessionId: string,
  editId: string,
  decision: "accept" | "reject"
): Promise<void> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/edits/${editId}/decision`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ decision })
  });
  await parseJson(response);
}

export async function promoteWorking(sessionId: string): Promise<void> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/promote-working`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ confirm: true })
  });
  await parseJson(response);
}

export function workingDownloadUrl(sessionId: string): string {
  return `${API_BASE}/api/session/${sessionId}/download?variant=working`;
}


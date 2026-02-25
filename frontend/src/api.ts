import { ProposalBatch, SessionState } from "./types";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:8080";

export type Provider = "anthropic" | "gemini" | "openrouter";

async function parseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    let detail = `Request failed (${response.status})`;
    try {
      const body = (await response.json()) as { error?: string };
      if (body.error) {
        detail = body.error;
      }
    } catch {
      try {
        const raw = await response.text();
        if (raw.includes("Cannot POST")) {
          detail = "Endpoint is unavailable. Restart the backend container and retry.";
        }
      } catch {
        // Keep default message.
      }
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

export async function restoreSavedSession(): Promise<
  | { id: string; createdAt: string; restored: true; savedAt: string }
  | null
> {
  const response = await fetch(`${API_BASE}/api/session/restore-saved`, {
    method: "POST"
  });
  if (response.status === 404) {
    return null;
  }
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

export async function removeSource(sessionId: string): Promise<void> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/source`, {
    method: "DELETE"
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

export async function removeContext(sessionId: string, contextId: string): Promise<void> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/context/${contextId}`, {
    method: "DELETE"
  });
  await parseJson(response);
}

export async function applyManualEdit(
  sessionId: string,
  edits: Array<{ blockId: string; text: string }>
): Promise<{ updatedCount: number }> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/manual-edit`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ edits })
  });
  return parseJson(response);
}

export async function proposeEdits(
  sessionId: string,
  payload: {
    prompt: string;
    provider: Provider;
    model?: string;
    apiKey: string;
  }
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

export async function analyzeGrammar(
  sessionId: string,
  payload: {
    customInstructions?: string;
    provider: Provider;
    model?: string;
    apiKey: string;
  }
): Promise<ProposalBatch> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/analyze-grammar`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  return parseJson(response);
}

export async function fetchProviderModels(payload: {
  provider: Provider;
  apiKey: string;
}): Promise<{
  provider: Provider;
  defaultModel: string;
  models: Array<{ id: string; label?: string }>;
}> {
  const response = await fetch(`${API_BASE}/api/models`, {
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
  decision: "accept" | "reject",
  wordChangeIndex?: number
): Promise<void> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/edits/${editId}/decision`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      decision,
      ...(typeof wordChangeIndex === "number" ? { wordChangeIndex } : {})
    })
  });
  await parseJson(response);
}

export async function saveWorkspace(sessionId: string): Promise<{ savedAt: string }> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/save-workspace`, {
    method: "POST"
  });
  return parseJson(response);
}

export async function removeSavedWorkspace(): Promise<{ removed: boolean }> {
  const response = await fetch(`${API_BASE}/api/session/saved-workspace`, {
    method: "DELETE"
  });
  return parseJson(response);
}

export async function acceptAllEdits(sessionId: string): Promise<{ acceptedCount: number }> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/edits/accept-all`, {
    method: "POST"
  });
  return parseJson(response);
}

export function workingDownloadUrl(sessionId: string): string {
  return `${API_BASE}/api/session/${sessionId}/download?variant=working`;
}

export async function fetchPreviewDoc(
  sessionId: string,
  variant: "working" | "source" = "working"
): Promise<ArrayBuffer> {
  const response = await fetch(`${API_BASE}/api/session/${sessionId}/preview?variant=${variant}`, {
    method: "GET",
    cache: "no-store"
  });
  if (!response.ok) {
    throw new Error("Failed to load formatted document preview.");
  }
  return response.arrayBuffer();
}

import { ChangeEvent, useEffect, useMemo, useState } from "react";
import {
  createSession,
  decideEdit,
  fetchState,
  promoteWorking,
  proposeEdits,
  uploadContext,
  uploadSource,
  workingDownloadUrl
} from "./api";
import { ProposalBatch, SessionState } from "./types";

type Provider = "anthropic" | "gemini" | "openrouter" | "mock";

function formatDate(iso: string): string {
  return new Date(iso).toLocaleString();
}

function App() {
  const [sessionId, setSessionId] = useState<string>("");
  const [state, setState] = useState<SessionState | null>(null);
  const [prompt, setPrompt] = useState("");
  const [provider, setProvider] = useState<Provider>("mock");
  const [model, setModel] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("Creating session...");

  async function refresh(targetSessionId: string): Promise<void> {
    const next = await fetchState(targetSessionId);
    setState(next);
  }

  useEffect(() => {
    async function init() {
      try {
        setLoading(true);
        const session = await createSession();
        setSessionId(session.id);
        await refresh(session.id);
        setStatus(`Session ${session.id.slice(0, 8)} ready.`);
      } catch (initError) {
        setError(initError instanceof Error ? initError.message : "Failed to create session.");
      } finally {
        setLoading(false);
      }
    }
    init();
  }, []);

  async function onSourceFileChange(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    if (!sessionId) {
      return;
    }
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) {
      return;
    }

    try {
      setLoading(true);
      setError("");
      await uploadSource(sessionId, file);
      await refresh(sessionId);
      setStatus(`Loaded source document: ${file.name}`);
    } catch (uploadError) {
      setError(uploadError instanceof Error ? uploadError.message : "Failed to upload document.");
    } finally {
      setLoading(false);
    }
  }

  async function onContextFileChange(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    if (!sessionId) {
      return;
    }
    const files = event.target.files ? Array.from(event.target.files) : [];
    event.target.value = "";
    if (files.length === 0) {
      return;
    }

    try {
      setLoading(true);
      setError("");
      for (const file of files) {
        await uploadContext(sessionId, file);
      }
      await refresh(sessionId);
      setStatus(`Uploaded ${files.length} context file(s).`);
    } catch (uploadError) {
      setError(uploadError instanceof Error ? uploadError.message : "Failed to upload context.");
    } finally {
      setLoading(false);
    }
  }

  async function onPropose(): Promise<void> {
    if (!sessionId || !prompt.trim()) {
      return;
    }

    try {
      setLoading(true);
      setError("");
      const payload = {
        prompt: prompt.trim(),
        provider,
        ...(model.trim() ? { model: model.trim() } : {})
      };
      const result = await proposeEdits(sessionId, payload);
      await refresh(sessionId);
      setStatus(`Generated ${result.edits.length} proposal(s) via ${result.provider}.`);
    } catch (proposeError) {
      setError(proposeError instanceof Error ? proposeError.message : "Failed to generate edits.");
    } finally {
      setLoading(false);
    }
  }

  async function onDecide(editId: string, decision: "accept" | "reject"): Promise<void> {
    if (!sessionId) {
      return;
    }
    try {
      setLoading(true);
      setError("");
      await decideEdit(sessionId, editId, decision);
      await refresh(sessionId);
      setStatus(`Edit ${decision}ed.`);
    } catch (decisionError) {
      setError(decisionError instanceof Error ? decisionError.message : "Failed to apply decision.");
    } finally {
      setLoading(false);
    }
  }

  async function onPromoteWorking(): Promise<void> {
    if (!sessionId) {
      return;
    }
    const confirmed = window.confirm(
      "This will overwrite the session's source copy with the current working document. Continue?"
    );
    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);
      setError("");
      await promoteWorking(sessionId);
      await refresh(sessionId);
      setStatus("Working copy promoted to source baseline.");
    } catch (promoteError) {
      setError(promoteError instanceof Error ? promoteError.message : "Failed to promote working copy.");
    } finally {
      setLoading(false);
    }
  }

  const allBatches = useMemo(() => {
    if (!state) {
      return [] as ProposalBatch[];
    }
    return [...state.proposalHistory].reverse();
  }, [state]);

  const pendingCount = useMemo(() => {
    if (!state) {
      return 0;
    }
    return state.proposalHistory
      .flatMap((batch) => batch.edits)
      .filter((edit) => edit.status === "pending").length;
  }, [state]);

  return (
    <div className="app">
      <header className="topbar">
        <h1>Doc Editing Application</h1>
        <div className="meta">
          <span>{status}</span>
          <span>Pending edits: {pendingCount}</span>
        </div>
      </header>

      <section className="controls">
        <label className="file-control">
          Source `.docx`
          <input type="file" accept=".docx" onChange={onSourceFileChange} disabled={loading} />
        </label>

        <label className="file-control">
          Context Files (`.docx`, `.txt`, `.md`)
          <input
            type="file"
            accept=".docx,.txt,.md"
            multiple
            onChange={onContextFileChange}
            disabled={loading}
          />
        </label>

        <div className="provider-row">
          <label>
            Provider
            <select
              value={provider}
              onChange={(event) => setProvider(event.target.value as Provider)}
              disabled={loading}
            >
              <option value="mock">Mock (no API key required)</option>
              <option value="anthropic">Anthropic (Claude)</option>
              <option value="gemini">Gemini</option>
              <option value="openrouter">OpenRouter</option>
            </select>
          </label>
          <label>
            Model override (optional)
            <input
              value={model}
              onChange={(event) => setModel(event.target.value)}
              placeholder="Leave blank for default"
              disabled={loading}
            />
          </label>
          {sessionId && (
            <a
              className="download-link"
              href={workingDownloadUrl(sessionId)}
              target="_blank"
              rel="noreferrer"
            >
              Download Working `.docx`
            </a>
          )}
          <button type="button" onClick={onPromoteWorking} disabled={loading || !state?.workingBlocks.length}>
            Promote Working to Source
          </button>
        </div>

        <label className="prompt-field">
          Edit instruction
          <textarea
            value={prompt}
            onChange={(event) => setPrompt(event.target.value)}
            placeholder="Example: tighten wording and fix punctuation for executive tone."
            rows={4}
            disabled={loading}
          />
        </label>
        <button type="button" onClick={onPropose} disabled={loading || !state?.workingBlocks.length}>
          Propose Edits
        </button>
        {error && <p className="error">{error}</p>}
      </section>

      <main className="workspace">
        <section className="panel">
          <h2>Current Working Document</h2>
          {state?.sourceFilename && <p className="subtle">Source file: {state.sourceFilename}</p>}
          {!state?.workingBlocks.length && (
            <p className="empty">Upload a `.docx` source document to start editing.</p>
          )}
          <div className="doc-view">
            {state?.workingBlocks.map((block, idx) => (
              <p key={block.id} className="doc-block">
                <span className="block-number">{idx + 1}.</span> {block.text}
              </p>
            ))}
          </div>
          {!!state?.contextFiles.length && (
            <div className="context-list">
              <h3>Context Files</h3>
              {state.contextFiles.map((file) => (
                <p key={file.id}>
                  {file.filename} ({file.charCount.toLocaleString()} chars)
                </p>
              ))}
            </div>
          )}
        </section>

        <section className="panel">
          <h2>Proposed Edits</h2>
          {!allBatches.length && (
            <p className="empty">Run a prompt to generate edit proposals for review.</p>
          )}
          <div className="proposal-list">
            {allBatches.map((batch) => (
              <article key={batch.id} className="batch">
                <header className="batch-header">
                  <p>
                    {formatDate(batch.createdAt)} | {batch.provider} | {batch.model}
                  </p>
                  <p className="subtle">Prompt: {batch.prompt}</p>
                </header>
                {!batch.edits.length && <p className="empty">No changes were proposed for this prompt.</p>}
                {batch.edits.map((edit) => (
                  <div key={edit.id} className={`edit-card status-${edit.status}`}>
                    <div
                      className="diff-content"
                      dangerouslySetInnerHTML={{
                        __html: edit.diffHtml
                      }}
                    />
                    <p className="rationale">{edit.rationale}</p>
                    <div className="actions">
                      <span className="badge">{edit.status}</span>
                      <button
                        type="button"
                        disabled={loading || edit.status !== "pending"}
                        onClick={() => onDecide(edit.id, "accept")}
                      >
                        Accept
                      </button>
                      <button
                        type="button"
                        disabled={loading || edit.status !== "pending"}
                        onClick={() => onDecide(edit.id, "reject")}
                      >
                        Reject
                      </button>
                    </div>
                  </div>
                ))}
              </article>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
}

export default App;

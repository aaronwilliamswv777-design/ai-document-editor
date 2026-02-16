import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import { renderAsync } from "docx-preview";
import {
  acceptAllEdits,
  analyzeGrammar,
  createSession,
  decideEdit,
  fetchProviderModels,
  fetchPreviewDoc,
  fetchState,
  Provider,
  promoteWorking,
  proposeEdits,
  uploadContext,
  uploadSource,
  workingDownloadUrl
} from "./api";
import { ProposalBatch, SessionState } from "./types";

type EditMode = "custom" | "grammar";
type ProviderKeyMap = Record<Provider, string>;
type ProviderModelMap = Record<Provider, Array<{ id: string; label?: string }>>;

type GrammarHighlight = {
  blockId: string;
  blockText: string;
  targetText: string;
  tooltip: string;
  hintIndex: number;
};

function formatDate(iso: string): string {
  return new Date(iso).toLocaleString();
}

function isInsideExistingGrammarHighlight(node: Node): boolean {
  let current: Node | null = node.parentNode;
  while (current) {
    if (current.nodeType === Node.ELEMENT_NODE) {
      const element = current as Element;
      if (
        element.classList.contains("grammar-issue-highlight") ||
        element.classList.contains("grammar-issue-highlight-whole")
      ) {
        return true;
      }
    }
    current = current.parentNode;
  }
  return false;
}

function collectTextNodes(element: Element): Text[] {
  const nodes: Text[] = [];
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let current: Node | null = walker.nextNode();
  while (current) {
    if (!isInsideExistingGrammarHighlight(current)) {
      nodes.push(current as Text);
    }
    current = walker.nextNode();
  }
  return nodes;
}

function normalizeSearchText(value: string): string {
  return value
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function extractAtomicHighlightTargets(value: string): string[] {
  const matches = value.match(/[A-Za-z0-9']+|[.,;:!?]/g);
  if (!matches) {
    return [];
  }
  return matches.map((item) => item.trim()).filter(Boolean);
}

function isWordToken(value: string): boolean {
  return /^[A-Za-z0-9']+$/.test(value);
}

function isWordChar(value: string): boolean {
  return /[A-Za-z0-9']/.test(value);
}

function findNeedleStart(fullText: string, needle: string): number {
  const loweredText = fullText.toLowerCase();
  const loweredNeedle = needle.toLowerCase();
  let cursor = 0;

  while (cursor <= loweredText.length - loweredNeedle.length) {
    const index = loweredText.indexOf(loweredNeedle, cursor);
    if (index < 0) {
      return -1;
    }

    if (!isWordToken(needle)) {
      return index;
    }

    const prev = index > 0 ? fullText[index - 1] : "";
    const next = index + needle.length < fullText.length ? fullText[index + needle.length] : "";
    const boundaryOk = (!prev || !isWordChar(prev)) && (!next || !isWordChar(next));
    if (boundaryOk) {
      return index;
    }
    cursor = index + 1;
  }

  return -1;
}

function splitIntoSentences(text: string): string[] {
  const normalized = text
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) {
    return [];
  }

  const matches = normalized.match(/[^.!?]+[.!?]?/g) || [];
  const sentences = matches.map((item) => item.trim()).filter(Boolean);
  return sentences.length > 0 ? sentences : [normalized];
}

function buildSentenceScopedSuggestion(proposedText: string, targetText: string): string {
  const sentences = splitIntoSentences(proposedText);
  if (sentences.length === 0) {
    return proposedText.trim();
  }

  const normalizedTarget = normalizeSearchText(targetText);
  if (normalizedTarget) {
    const directMatch = sentences.find((sentence) =>
      normalizeSearchText(sentence).includes(normalizedTarget)
    );
    if (directMatch) {
      return directMatch;
    }

    const targetTokens = normalizedTarget.split(/\s+/).filter(Boolean);
    for (const token of targetTokens) {
      if (token.length < 2) {
        continue;
      }
      const tokenMatch = sentences.find((sentence) =>
        normalizeSearchText(sentence).includes(token)
      );
      if (tokenMatch) {
        return tokenMatch;
      }
    }
  }

  return sentences[0];
}

function applyHighlightToParagraph(paragraph: Element, targetText: string, tooltip: string): boolean {
  const needle = targetText.trim();
  if (!needle) {
    return false;
  }

  const textNodes = collectTextNodes(paragraph).filter((node) => (node.nodeValue || "").length > 0);
  if (textNodes.length === 0) {
    return false;
  }

  const fullText = textNodes.map((node) => node.nodeValue || "").join("");
  const start = findNeedleStart(fullText, needle);
  if (start < 0) {
    return false;
  }
  const end = start + needle.length;

  let cursor = 0;
  let startNode: Text | null = null;
  let endNode: Text | null = null;
  let startOffset = 0;
  let endOffset = 0;

  for (const node of textNodes) {
    const value = node.nodeValue || "";
    const nextCursor = cursor + value.length;

    if (!startNode && start >= cursor && start <= nextCursor) {
      startNode = node;
      startOffset = Math.max(0, start - cursor);
    }
    if (!endNode && end >= cursor && end <= nextCursor) {
      endNode = node;
      endOffset = Math.max(0, end - cursor);
    }

    cursor = nextCursor;
    if (startNode && endNode) {
      break;
    }
  }

  if (!startNode || !endNode) {
    return false;
  }

  const range = document.createRange();
  range.setStart(startNode, startOffset);
  range.setEnd(endNode, endOffset);

  const marker = document.createElement("span");
  marker.className = "grammar-issue-highlight";
  marker.setAttribute("data-grammar-tooltip", tooltip);
  marker.setAttribute("aria-label", tooltip);

  try {
    range.surroundContents(marker);
  } catch {
    const extracted = range.extractContents();
    marker.appendChild(extracted);
    range.insertNode(marker);
  }

  return true;
}

function applyGrammarHighlights(container: HTMLElement, highlights: GrammarHighlight[]): void {
  if (highlights.length === 0) {
    return;
  }

  const paragraphs = Array.from(container.querySelectorAll(".docx p, .docx li, .docx td, .docx th"));
  const candidateParagraphs = paragraphs.length
    ? paragraphs
    : Array.from(container.querySelectorAll(".docx *")).filter(
        (element) => normalizeSearchText(element.textContent || "").length > 0
      );

  for (const highlight of highlights) {
    let paragraph: Element | undefined;
    const normalizedBlock = normalizeSearchText(highlight.blockText);
    if (normalizedBlock) {
      paragraph = candidateParagraphs.find((item) =>
        normalizeSearchText(item.textContent || "").includes(normalizedBlock.slice(0, 120))
      );
    }
    if (!paragraph) {
      paragraph = candidateParagraphs[highlight.hintIndex];
    }
    if (!paragraph) {
      paragraph = candidateParagraphs.find((item) =>
        normalizeSearchText(item.textContent || "").includes(normalizeSearchText(highlight.targetText))
      );
    }
    if (!paragraph) {
      continue;
    }

    applyHighlightToParagraph(paragraph, highlight.targetText, highlight.tooltip);
  }
}

function App() {
  const uiRevision = "model-picker-v3";
  const [sessionId, setSessionId] = useState<string>("");
  const [state, setState] = useState<SessionState | null>(null);
  const [instructionText, setInstructionText] = useState("");
  const [editMode, setEditMode] = useState<EditMode>("custom");
  const [provider, setProvider] = useState<Provider>("anthropic");
  const [model, setModel] = useState("");
  const [providerModels, setProviderModels] = useState<ProviderModelMap>({
    anthropic: [],
    gemini: [],
    openrouter: []
  });
  const [showApiKeyMenu, setShowApiKeyMenu] = useState(false);
  const [apiKeys, setApiKeys] = useState<ProviderKeyMap>({
    anthropic: "",
    gemini: "",
    openrouter: ""
  });
  const [loading, setLoading] = useState(false);
  const [loadingModels, setLoadingModels] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("Creating session...");
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState("");
  const previewContainerRef = useRef<HTMLDivElement | null>(null);

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

  async function onRunModeAction(): Promise<void> {
    if (!sessionId || !state?.workingBlocks.length) {
      return;
    }
    if (editMode === "custom" && !instructionText.trim()) {
      return;
    }
    const selectedApiKey = apiKeys[provider].trim();
    if (!selectedApiKey) {
      setError(`Add a ${provider} API key in the API Keys menu before running this mode.`);
      return;
    }

    try {
      setLoading(true);
      setError("");

      if (editMode === "grammar") {
        const result = await analyzeGrammar(sessionId, {
          customInstructions: instructionText.trim() || undefined,
          provider,
          apiKey: selectedApiKey,
          ...(model.trim() ? { model: model.trim() } : {})
        });
        await refresh(sessionId);
        setStatus(
          `Detected ${result.edits.length} grammar/punctuation issue(s) via ${result.provider}.`
        );
      } else {
        const result = await proposeEdits(sessionId, {
          prompt: instructionText.trim(),
          provider,
          apiKey: selectedApiKey,
          ...(model.trim() ? { model: model.trim() } : {})
        });
        await refresh(sessionId);
        setStatus(`Generated ${result.edits.length} proposal(s) via ${result.provider}.`);
      }
    } catch (runError) {
      setError(runError instanceof Error ? runError.message : "Failed to run analysis.");
    } finally {
      setLoading(false);
    }
  }

  async function onLoadModels(): Promise<void> {
    const selectedApiKey = apiKeys[provider].trim();
    if (!selectedApiKey) {
      setError(`Add a ${provider} API key first, then load models.`);
      return;
    }

    try {
      setLoadingModels(true);
      setError("");
      const result = await fetchProviderModels({
        provider,
        apiKey: selectedApiKey
      });

      setProviderModels((prev) => ({
        ...prev,
        [provider]: result.models
      }));

      if (result.models.length === 0) {
        setModel("");
        setStatus(`No models returned for ${provider}.`);
        return;
      }

      const hasCurrent = result.models.some((item) => item.id === model);
      if (!hasCurrent && !model.trim()) {
        const defaultExists = result.models.some((item) => item.id === result.defaultModel);
        const nextModel = defaultExists ? result.defaultModel : result.models[0].id;
        setModel(nextModel);
      }

      setStatus(`Loaded ${result.models.length} models for ${provider}.`);
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : "Failed to load models.");
    } finally {
      setLoadingModels(false);
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

  async function onAcceptAllPending(): Promise<void> {
    if (!sessionId) {
      return;
    }
    try {
      setLoading(true);
      setError("");
      const result = await acceptAllEdits(sessionId);
      await refresh(sessionId);
      setStatus(`Accepted ${result.acceptedCount} pending edit(s).`);
    } catch (acceptError) {
      setError(acceptError instanceof Error ? acceptError.message : "Failed to accept all edits.");
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

  const filteredProviderModels = useMemo(() => {
    const query = model.trim().toLowerCase();
    const options = providerModels[provider];
    if (!query) {
      return options;
    }
    return options.filter((item) =>
      `${item.label || ""} ${item.id}`.toLowerCase().includes(query)
    );
  }, [model, provider, providerModels]);

  const selectedKnownModelId = useMemo(() => {
    const known = providerModels[provider].some((item) => item.id === model.trim());
    return known ? model.trim() : "";
  }, [providerModels, provider, model]);

  const grammarHighlights = useMemo(() => {
    if (!state) {
      return [] as GrammarHighlight[];
    }

    const blockMap = new Map(state.workingBlocks.map((block) => [block.id, block]));
    const blockIndexById = new Map(state.workingBlocks.map((block, index) => [block.id, index]));
    const highlights: GrammarHighlight[] = [];
    const seen = new Set<string>();

    for (const batch of state.proposalHistory) {
      if (batch.mode !== "grammar") {
        continue;
      }
      for (const edit of batch.edits) {
        if (edit.status !== "pending") {
          continue;
        }
        const block = blockMap.get(edit.blockId);
        if (!block) {
          continue;
        }
        const rawTargetTexts = edit.highlightTexts?.length
          ? edit.highlightTexts
          : [edit.highlightText || edit.originalText];

        const targetTexts = Array.from(
          new Set(
            rawTargetTexts
              .flatMap((value) => extractAtomicHighlightTargets(value))
              .map((value) => value.trim())
              .filter(Boolean)
          )
        ).slice(0, 12);

        if (targetTexts.length === 0) {
          continue;
        }

        for (const targetText of targetTexts) {
          const key = `${edit.blockId}:${targetText.toLowerCase()}`;
          if (seen.has(key)) {
            continue;
          }
          seen.add(key);
          highlights.push({
            blockId: edit.blockId,
            blockText: block.text,
            targetText,
            tooltip: `Reason\n${edit.rationale}\n\nSuggested change\n${buildSentenceScopedSuggestion(
              edit.proposedText,
              targetText
            )}`,
            hintIndex: blockIndexById.get(edit.blockId) ?? 0
          });
        }
      }
    }

    return highlights;
  }, [state]);

  useEffect(() => {
    let cancelled = false;

    async function renderDocPreview() {
      const container = previewContainerRef.current;
      if (!sessionId || !state?.workingBlocks.length || !container) {
        if (container) {
          container.innerHTML = "";
        }
        setPreviewLoading(false);
        setPreviewError("");
        return;
      }

      try {
        setPreviewLoading(true);
        setPreviewError("");
        const fileBuffer = await fetchPreviewDoc(sessionId, "working");
        if (cancelled || !previewContainerRef.current) {
          return;
        }
        previewContainerRef.current.innerHTML = "";
        await renderAsync(fileBuffer, previewContainerRef.current);
        applyGrammarHighlights(previewContainerRef.current, grammarHighlights);
      } catch (renderError) {
        if (!cancelled) {
          setPreviewError(
            renderError instanceof Error
              ? renderError.message
              : "Failed to render the Word preview."
          );
        }
      } finally {
        if (!cancelled) {
          setPreviewLoading(false);
        }
      }
    }

    renderDocPreview();
    return () => {
      cancelled = true;
    };
  }, [sessionId, state, grammarHighlights]);

  return (
    <div className="app">
      <header className="topbar">
        <h1>Doc Editing Application</h1>
        <div className="meta">
          <span>{status}</span>
          <span>Pending edits: {pendingCount}</span>
          <span>UI rev: {uiRevision}</span>
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
              onChange={(event) => {
                setProvider(event.target.value as Provider);
                setModel("");
              }}
              disabled={loading || loadingModels}
            >
              <option value="anthropic">Anthropic (Claude)</option>
              <option value="gemini">Gemini</option>
              <option value="openrouter">OpenRouter</option>
            </select>
          </label>
          <div className="model-picker-block">
            <span className="mode-label">Model Picker</span>
            <div className="model-picker-layout">
              <label className="model-available-column">
                Available models (left)
                <select
                  className="model-available-select"
                  size={8}
                  value={selectedKnownModelId}
                  onChange={(event) => setModel(event.target.value)}
                  disabled={loading || loadingModels || providerModels[provider].length === 0}
                >
                  <option value="" disabled>
                    {providerModels[provider].length === 0
                      ? "Load models first"
                      : "Type on the right to filter, or click a model"}
                  </option>
                  {filteredProviderModels.length === 0 &&
                    providerModels[provider].length > 0 && (
                      <option value="__no_match__" disabled>
                        No models match the text on the right
                      </option>
                    )}
                  {filteredProviderModels.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.label ? `${item.label} (${item.id})` : item.id}
                    </option>
                  ))}
                </select>
              </label>
              <label className="model-typed-column">
                Model override (right)
                <input
                  type="text"
                  className="model-search-input"
                  value={model}
                  onChange={(event) => setModel(event.target.value)}
                  placeholder="Type model id here (or click one on the left)"
                  autoComplete="off"
                  spellCheck={false}
                  disabled={loading || loadingModels}
                />
                <p className="subtle model-override-readout">
                  Current override: {model.trim() || "Use provider default"}
                </p>
              </label>
            </div>
          </div>
          <p className="subtle">
            Showing {filteredProviderModels.length} of {providerModels[provider].length} models
          </p>
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

        <div className="api-key-menu">
          <button
            type="button"
            className={`api-key-toggle ${showApiKeyMenu ? "api-key-toggle-active" : ""}`}
            onClick={() => setShowApiKeyMenu((prev) => !prev)}
            disabled={loading || loadingModels}
          >
            {showApiKeyMenu ? "Hide API Key Menu" : "API Key Menu"}
          </button>
          <p className="subtle">
            Active provider key: {apiKeys[provider].trim() ? "Configured" : "Missing"}
          </p>
          <button type="button" className="api-key-load-btn" onClick={onLoadModels} disabled={loading || loadingModels}>
            {loadingModels ? "Loading Models..." : `Load Models For ${provider}`}
          </button>
          {showApiKeyMenu && (
            <div className="api-key-panel">
              <label>
                Anthropic API Key
                <input
                  type="password"
                  value={apiKeys.anthropic}
                  onChange={(event) =>
                    setApiKeys((prev) => ({
                      ...prev,
                      anthropic: event.target.value
                    }))
                  }
                  placeholder="sk-ant-..."
                  autoComplete="off"
                  disabled={loading || loadingModels}
                />
              </label>
              <label>
                Gemini API Key
                <input
                  type="password"
                  value={apiKeys.gemini}
                  onChange={(event) =>
                    setApiKeys((prev) => ({
                      ...prev,
                      gemini: event.target.value
                    }))
                  }
                  placeholder="AIza..."
                  autoComplete="off"
                  disabled={loading || loadingModels}
                />
              </label>
              <label>
                OpenRouter API Key
                <input
                  type="password"
                  value={apiKeys.openrouter}
                  onChange={(event) =>
                    setApiKeys((prev) => ({
                      ...prev,
                      openrouter: event.target.value
                    }))
                  }
                  placeholder="sk-or-v1-..."
                  autoComplete="off"
                  disabled={loading || loadingModels}
                />
              </label>
              <p className="subtle">Keys are used only for this browser session.</p>
            </div>
          )}
        </div>

        <div className="mode-row">
          <span className="mode-label">Mode</span>
          <button
            type="button"
            className={`mode-button ${editMode === "custom" ? "mode-button-active" : ""}`}
            onClick={() => setEditMode("custom")}
            disabled={loading}
          >
            Targeted Edit
          </button>
          <button
            type="button"
            className={`mode-button ${editMode === "grammar" ? "mode-button-active" : ""}`}
            onClick={() => setEditMode("grammar")}
            disabled={loading}
          >
            Grammar & Punctuation
          </button>
        </div>

        <label className="prompt-field">
          {editMode === "custom"
            ? "Edit instruction"
            : "Custom instructions for grammar mode (optional)"}
          <textarea
            value={instructionText}
            onChange={(event) => setInstructionText(event.target.value)}
            placeholder={
              editMode === "custom"
                ? "Example: tighten wording and fix punctuation for executive tone."
                : "Optional: focus on commas, tense consistency, and business formal grammar."
            }
            rows={4}
            disabled={loading}
          />
        </label>
        <button
          type="button"
          onClick={onRunModeAction}
          disabled={
            loading ||
            !state?.workingBlocks.length ||
            (editMode === "custom" && !instructionText.trim())
          }
        >
          {editMode === "custom" ? "Propose Edits" : "Analyze Grammar & Punctuation"}
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
          {previewLoading && <p className="subtle">Rendering formatted preview...</p>}
          {previewError && <p className="error">{previewError}</p>}
          <div className="doc-view doc-view-formatted">
            <div ref={previewContainerRef} className="docx-host" />
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
          <div className="panel-title-row">
            <h2>Proposed Edits</h2>
            <button type="button" onClick={onAcceptAllPending} disabled={loading || pendingCount === 0}>
              Accept All Pending
            </button>
          </div>
          {!allBatches.length && (
            <p className="empty">Run a prompt or grammar analysis to generate proposals for review.</p>
          )}
          <div className="proposal-list">
            {allBatches.map((batch) => (
              <article key={batch.id} className="batch">
                <header className="batch-header">
                  <p>
                    {formatDate(batch.createdAt)} | {batch.provider} | {batch.model}
                  </p>
                  <p className="subtle">
                    Mode: {batch.mode === "grammar" ? "Grammar & Punctuation" : "Targeted Edit"}
                  </p>
                  <p className="subtle">Instruction: {batch.prompt}</p>
                </header>
                {!batch.edits.length && <p className="empty">No changes were proposed for this run.</p>}
                {batch.edits.map((edit) => (
                  <div key={edit.id} className={`edit-card status-${edit.status}`}>
                    <div
                      className="diff-content"
                      dangerouslySetInnerHTML={{
                        __html: edit.diffHtml
                      }}
                    />
                    {batch.mode === "grammar" &&
                      (edit.highlightTexts?.length || edit.highlightText) && (
                        <p className="subtle">
                          Detected issue text: "
                          {(
                            edit.highlightTexts?.length
                              ? edit.highlightTexts.slice(0, 4).join('", "')
                              : edit.highlightText || ""
                          )}
                          "
                        </p>
                      )}
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

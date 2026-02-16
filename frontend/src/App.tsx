import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import { renderAsync } from "docx-preview";
import {
  acceptAllEdits,
  applyManualEdit,
  analyzeGrammar,
  createSession,
  decideEdit,
  fetchProviderModels,
  fetchPreviewDoc,
  fetchState,
  Provider,
  proposeEdits,
  removeContext,
  removeSavedWorkspace,
  removeSource,
  restoreSavedSession,
  saveWorkspace,
  uploadContext,
  uploadSource,
  workingDownloadUrl
} from "./api";
import { ProposalBatch, SessionState } from "./types";

type EditMode = "custom" | "grammar";
type TopMenu = "editor" | "settings";
type ProviderKeyMap = Record<Provider, string>;
type ProviderModelMap = Record<Provider, Array<{ id: string; label?: string }>>;
type PersistedPreferences = {
  provider?: Provider;
  model?: string;
  apiKeys?: Partial<ProviderKeyMap>;
};
const PREFERENCES_STORAGE_KEY = "doc-edit.preferences.v1";
const REMEMBER_SETTINGS_KEY = "doc-edit.remember-settings.v1";
const MODEL_CATALOG_STORAGE_KEY = "doc-edit.model-catalog.v1";

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

function isProviderValue(value: unknown): value is Provider {
  return value === "anthropic" || value === "gemini" || value === "openrouter";
}

function createEmptyProviderModels(): ProviderModelMap {
  return {
    anthropic: [],
    gemini: [],
    openrouter: []
  };
}

function sanitizeProviderModels(value: unknown): ProviderModelMap {
  const fallback = createEmptyProviderModels();
  if (!value || typeof value !== "object") {
    return fallback;
  }

  const parseList = (raw: unknown): Array<{ id: string; label?: string }> => {
    if (!Array.isArray(raw)) {
      return [];
    }
    const seen = new Set<string>();
    const result: Array<{ id: string; label?: string }> = [];
    for (const item of raw) {
      if (!item || typeof item !== "object") {
        continue;
      }
      const maybe = item as { id?: unknown; label?: unknown };
      if (typeof maybe.id !== "string") {
        continue;
      }
      const id = maybe.id.trim();
      if (!id || seen.has(id)) {
        continue;
      }
      seen.add(id);
      if (typeof maybe.label === "string" && maybe.label.trim()) {
        result.push({ id, label: maybe.label.trim() });
      } else {
        result.push({ id });
      }
      if (result.length >= 1200) {
        break;
      }
    }
    return result;
  };

  const parsed = value as Partial<Record<Provider, unknown>>;
  return {
    anthropic: parseList(parsed.anthropic),
    gemini: parseList(parsed.gemini),
    openrouter: parseList(parsed.openrouter)
  };
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

function normalizeEditableText(value: string): string {
  return value
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/\n+/g, " ")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function collectEditableParagraphElements(container: HTMLElement): HTMLElement[] {
  const paragraphs = Array.from(
    container.querySelectorAll<HTMLElement>(".docx p, .docx li, .docx td, .docx th")
  );
  if (paragraphs.length > 0) {
    return paragraphs;
  }
  return Array.from(container.querySelectorAll<HTMLElement>(".docx *")).filter(
    (element) => normalizeSearchText(element.textContent || "").length > 0
  );
}

function bindEditablePreviewBlocks(
  container: HTMLElement,
  blocks: SessionState["workingBlocks"],
  editable: boolean
): void {
  const candidates = collectEditableParagraphElements(container);
  candidates.forEach((element) => {
    element.removeAttribute("data-block-id");
    element.classList.remove("manual-editable-block");
    element.contentEditable = "false";
    element.spellcheck = false;
  });

  let cursor = 0;
  const usedIndices = new Set<number>();
  for (const block of blocks) {
    const normalizedBlockText = normalizeEditableText(block.text);
    if (!normalizedBlockText) {
      continue;
    }

    let matchIndex = -1;
    for (let index = cursor; index < candidates.length; index += 1) {
      if (usedIndices.has(index)) {
        continue;
      }
      const candidateText = normalizeEditableText(candidates[index].textContent || "");
      if (!candidateText) {
        continue;
      }
      if (candidateText === normalizedBlockText) {
        matchIndex = index;
        break;
      }
    }

    if (matchIndex < 0) {
      const blockHead = normalizedBlockText.slice(0, 80);
      for (let index = cursor; index < candidates.length; index += 1) {
        if (usedIndices.has(index)) {
          continue;
        }
        const candidateText = normalizeEditableText(candidates[index].textContent || "");
        if (!candidateText) {
          continue;
        }
        const candidateHead = candidateText.slice(0, 80);
        if (
          (blockHead.length >= 3 && candidateText.includes(blockHead)) ||
          (candidateHead.length >= 3 && normalizedBlockText.includes(candidateHead))
        ) {
          matchIndex = index;
          break;
        }
      }
    }

    if (matchIndex < 0) {
      continue;
    }

    const element = candidates[matchIndex];
    element.dataset.blockId = block.id;
    usedIndices.add(matchIndex);
    cursor = matchIndex + 1;

    if (editable) {
      element.contentEditable = "true";
      element.spellcheck = true;
      element.classList.add("manual-editable-block");
    }
  }
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
  const uiRevision = "menu-split-v16";
  const [sessionId, setSessionId] = useState<string>("");
  const [state, setState] = useState<SessionState | null>(null);
  const [instructionText, setInstructionText] = useState("");
  const [editMode, setEditMode] = useState<EditMode>("custom");
  const [topMenu, setTopMenu] = useState<TopMenu>("editor");
  const [provider, setProvider] = useState<Provider>("anthropic");
  const [model, setModel] = useState("");
  const [providerModels, setProviderModels] = useState<ProviderModelMap>(createEmptyProviderModels());
  const [showApiKeyMenu, setShowApiKeyMenu] = useState(false);
  const [apiKeys, setApiKeys] = useState<ProviderKeyMap>({
    anthropic: "",
    gemini: "",
    openrouter: ""
  });
  const [rememberPreferences, setRememberPreferences] = useState(false);
  const [preferencesLoaded, setPreferencesLoaded] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loadingModels, setLoadingModels] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("Creating session...");
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState("");
  const [directEditMode, setDirectEditMode] = useState(false);
  const previewContainerRef = useRef<HTMLDivElement | null>(null);

  async function refresh(targetSessionId: string): Promise<void> {
    const next = await fetchState(targetSessionId);
    setState(next);
  }

  useEffect(() => {
    try {
      const rawCatalog = window.localStorage.getItem(MODEL_CATALOG_STORAGE_KEY);
      if (rawCatalog) {
        const parsedCatalog = JSON.parse(rawCatalog) as unknown;
        setProviderModels(sanitizeProviderModels(parsedCatalog));
      }

      const rememberRaw = window.localStorage.getItem(REMEMBER_SETTINGS_KEY);
      const shouldRemember = rememberRaw === "1";
      setRememberPreferences(shouldRemember);
      if (!shouldRemember) {
        return;
      }

      const raw = window.localStorage.getItem(PREFERENCES_STORAGE_KEY);
      if (!raw) {
        return;
      }
      const parsed = JSON.parse(raw) as PersistedPreferences;
      if (isProviderValue(parsed.provider)) {
        setProvider(parsed.provider);
      }
      if (typeof parsed.model === "string") {
        setModel(parsed.model);
      }
      if (parsed.apiKeys && typeof parsed.apiKeys === "object") {
        setApiKeys((prev) => ({
          anthropic:
            typeof parsed.apiKeys?.anthropic === "string"
              ? parsed.apiKeys.anthropic
              : prev.anthropic,
          gemini:
            typeof parsed.apiKeys?.gemini === "string" ? parsed.apiKeys.gemini : prev.gemini,
          openrouter:
            typeof parsed.apiKeys?.openrouter === "string"
              ? parsed.apiKeys.openrouter
              : prev.openrouter
        }));
      }
    } catch {
      // Ignore corrupt local storage.
    } finally {
      setPreferencesLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (!preferencesLoaded) {
      return;
    }
    try {
      window.localStorage.setItem(REMEMBER_SETTINGS_KEY, rememberPreferences ? "1" : "0");
      if (!rememberPreferences) {
        window.localStorage.removeItem(PREFERENCES_STORAGE_KEY);
        return;
      }

      const payload: PersistedPreferences = { provider, model, apiKeys };
      window.localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // Ignore browser storage failures.
    }
  }, [preferencesLoaded, rememberPreferences, provider, model, apiKeys]);

  useEffect(() => {
    if (!preferencesLoaded) {
      return;
    }
    try {
      window.localStorage.setItem(MODEL_CATALOG_STORAGE_KEY, JSON.stringify(providerModels));
    } catch {
      // Ignore browser storage failures.
    }
  }, [preferencesLoaded, providerModels]);

  useEffect(() => {
    async function init() {
      try {
        setLoading(true);
        const restored = await restoreSavedSession();
        if (restored) {
          setSessionId(restored.id);
          await refresh(restored.id);
          setStatus(`Restored saved workspace from ${formatDate(restored.savedAt)}.`);
          return;
        }

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

  function onClearSavedPreferences(): void {
    try {
      window.localStorage.removeItem(PREFERENCES_STORAGE_KEY);
      window.localStorage.removeItem(REMEMBER_SETTINGS_KEY);
      window.localStorage.removeItem(MODEL_CATALOG_STORAGE_KEY);
    } catch {
      // Ignore browser storage failures.
    }

    setRememberPreferences(false);
    setProvider("anthropic");
    setModel("");
    setProviderModels(createEmptyProviderModels());
    setApiKeys({
      anthropic: "",
      gemini: "",
      openrouter: ""
    });
    setStatus("Cleared saved provider, model override, API keys, and model lists for this browser.");
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

  async function onSaveWorkspaceForReturn(): Promise<void> {
    if (!sessionId || !state?.workingBlocks.length) {
      return;
    }
    try {
      setLoading(true);
      setError("");
      const result = await saveWorkspace(sessionId);
      setStatus(`Saved current working document for next return (${formatDate(result.savedAt)}).`);
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : "Failed to save workspace.");
    } finally {
      setLoading(false);
    }
  }

  async function onRemoveSavedWorkspaceForReturn(): Promise<void> {
    const confirmed = window.confirm(
      "Remove the saved workspace for next return? Your current in-memory session will stay open."
    );
    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);
      setError("");
      const result = await removeSavedWorkspace();
      setStatus(
        result.removed
          ? "Removed saved workspace for next return."
          : "No saved workspace was found to remove."
      );
    } catch (removeError) {
      setError(
        removeError instanceof Error ? removeError.message : "Failed to remove saved workspace."
      );
    } finally {
      setLoading(false);
    }
  }

  async function onRemoveSourceDocument(): Promise<void> {
    if (!sessionId) {
      return;
    }
    const confirmed = window.confirm(
      "Remove the current source and working document from this session? Context files will stay."
    );
    if (!confirmed) {
      return;
    }
    try {
      setLoading(true);
      setError("");
      await removeSource(sessionId);
      await refresh(sessionId);
      setStatus("Removed source/working document. Upload another `.docx` to continue.");
    } catch (removeError) {
      setError(
        removeError instanceof Error ? removeError.message : "Failed to remove source document."
      );
    } finally {
      setLoading(false);
    }
  }

  async function onRemoveContextFile(contextId: string): Promise<void> {
    if (!sessionId) {
      return;
    }
    try {
      setLoading(true);
      setError("");
      await removeContext(sessionId, contextId);
      await refresh(sessionId);
      setStatus("Removed context file.");
    } catch (removeError) {
      setError(
        removeError instanceof Error ? removeError.message : "Failed to remove context file."
      );
    } finally {
      setLoading(false);
    }
  }

  function onToggleDirectEditMode(): void {
    setDirectEditMode((previous) => {
      const next = !previous;
      setStatus(
        next
          ? "Direct edit mode enabled. Click into the document preview and type."
          : "Direct edit mode disabled."
      );
      return next;
    });
  }

  function collectDirectPreviewEdits(): Array<{ blockId: string; text: string }> {
    if (!state?.workingBlocks.length || !previewContainerRef.current) {
      return [];
    }

    const blockTextById = new Map(state.workingBlocks.map((block) => [block.id, block.text]));
    const editableElements = Array.from(
      previewContainerRef.current.querySelectorAll<HTMLElement>("[data-block-id]")
    );

    const edits: Array<{ blockId: string; text: string }> = [];
    for (const element of editableElements) {
      const blockId = element.dataset.blockId;
      if (!blockId) {
        continue;
      }
      const currentText = blockTextById.get(blockId);
      if (typeof currentText !== "string") {
        continue;
      }

      const nextText = (element.textContent || "")
        .replace(/\u00a0/g, " ")
        .replace(/\r/g, "")
        .replace(/\n+/g, " ");
      if (nextText !== currentText) {
        edits.push({
          blockId,
          text: nextText
        });
      }
    }
    return edits;
  }

  async function onApplyDirectTextEdits(): Promise<void> {
    if (!sessionId || !state?.workingBlocks.length) {
      return;
    }

    const edits = collectDirectPreviewEdits();
    if (edits.length === 0) {
      setStatus("No typed changes to apply.");
      return;
    }

    try {
      setLoading(true);
      setError("");
      const result = await applyManualEdit(sessionId, edits);
      await refresh(sessionId);
      setStatus(`Applied ${result.updatedCount} typed change(s) to the working document.`);
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Failed to apply typed changes.");
    } finally {
      setLoading(false);
    }
  }

  async function onDiscardDirectTextEdits(): Promise<void> {
    if (!sessionId) {
      return;
    }
    try {
      setLoading(true);
      setError("");
      await refresh(sessionId);
      setStatus("Discarded unapplied direct edits.");
    } catch (discardError) {
      setError(
        discardError instanceof Error ? discardError.message : "Failed to discard direct edits."
      );
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
      if (topMenu !== "editor" || !sessionId || !state?.workingBlocks.length || !container) {
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
        if (!directEditMode) {
          applyGrammarHighlights(previewContainerRef.current, grammarHighlights);
        }
        bindEditablePreviewBlocks(previewContainerRef.current, state.workingBlocks, directEditMode);
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
  }, [topMenu, sessionId, state, grammarHighlights, directEditMode]);

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

      <section className="top-menu">
        <div className="top-menu-buttons">
          <button
            type="button"
            className={`top-menu-button ${topMenu === "editor" ? "top-menu-button-active" : ""}`}
            onClick={() => setTopMenu("editor")}
            disabled={loading || loadingModels}
          >
            Editor
          </button>
          <button
            type="button"
            className={`top-menu-button ${topMenu === "settings" ? "top-menu-button-active" : ""}`}
            onClick={() => setTopMenu("settings")}
            disabled={loading || loadingModels}
          >
            AI Settings
          </button>
        </div>
        <p className="subtle">
          Mode: {editMode === "grammar" ? "Grammar & Punctuation" : "Targeted Edit"} | Provider:{" "}
          {provider} | Model: {model.trim() || "Default"}
        </p>
      </section>

      {topMenu === "editor" ? (
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

          <div className="editor-mode-summary">
            <p className="subtle">
              Editing with{" "}
              {editMode === "grammar" ? "Grammar & Punctuation mode" : "Targeted Edit mode"}.
            </p>
            <button
              type="button"
              className="settings-jump-btn"
              onClick={() => setTopMenu("settings")}
              disabled={loading || loadingModels}
            >
              Open AI Settings
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
        </section>
      ) : (
        <section className="controls settings-controls">
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
            <label className="remember-toggle">
              <input
                type="checkbox"
                checked={rememberPreferences}
                onChange={(event) => setRememberPreferences(event.target.checked)}
                disabled={loading || loadingModels}
              />
              Remember keys + model on this device
            </label>
            <button type="button" className="api-key-load-btn" onClick={onLoadModels} disabled={loading || loadingModels}>
              {loadingModels ? "Loading Models..." : `Load Models For ${provider}`}
            </button>
            <button
              type="button"
              className="api-key-clear-btn"
              onClick={onClearSavedPreferences}
              disabled={loading || loadingModels}
            >
              Clear Saved Keys + Model
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
                <p className="subtle">
                  {rememberPreferences
                    ? "Keys, selected provider, and model override are saved in this browser."
                    : "Remember is off. Keys and model override are not saved after exit/reload."}
                </p>
              </div>
            )}
          </div>
        </section>
      )}
      {error && <p className="error">{error}</p>}

      {topMenu === "editor" && (
      <main className="workspace">
        <section className="panel">
          <div className="panel-title-row">
            <h2>Current Working Document</h2>
            <div className="panel-actions">
              <button
                type="button"
                className="session-save-btn"
                onClick={onSaveWorkspaceForReturn}
                disabled={loading || !state?.workingBlocks.length}
              >
                Save Working for Next Return
              </button>
              <button
                type="button"
                className="session-remove-saved-btn"
                onClick={onRemoveSavedWorkspaceForReturn}
                disabled={loading}
              >
                Remove Saved Workspace for Next Return
              </button>
              {sessionId && state?.workingBlocks.length ? (
                <a
                  className="download-link"
                  href={workingDownloadUrl(sessionId)}
                  target="_blank"
                  rel="noreferrer"
                >
                  Download Finished DOCX
                </a>
              ) : null}
              <button
                type="button"
                className="danger-btn"
                onClick={onRemoveSourceDocument}
                disabled={loading || !state?.workingBlocks.length}
              >
                Remove Current Document
              </button>
            </div>
          </div>
          {state?.sourceFilename && <p className="subtle">Source file: {state.sourceFilename}</p>}
          {!state?.workingBlocks.length && (
            <p className="empty">Upload a `.docx` source document to start editing.</p>
          )}
          {previewLoading && <p className="subtle">Rendering formatted preview...</p>}
          {previewError && <p className="error">{previewError}</p>}
          <div className="doc-view doc-view-formatted">
            <div ref={previewContainerRef} className="docx-host" />
          </div>
          {!!state?.workingBlocks.length && (
            <div className="manual-edit-panel">
              <div
                className={`manual-edit-header ${directEditMode ? "manual-edit-header-active" : ""}`}
              >
                <h3>Direct In-Document Editing</h3>
                <p className="subtle">
                  Turn this on, click directly in the preview document, and type like Word.
                </p>
              </div>
              <div className="manual-edit-actions">
                <button
                  type="button"
                  className={`direct-edit-toggle-btn ${directEditMode ? "direct-edit-toggle-btn-active" : ""}`}
                  onClick={onToggleDirectEditMode}
                  disabled={loading}
                >
                  {directEditMode ? "Disable Direct Typing" : "Enable Direct Typing"}
                </button>
                <button
                  type="button"
                  onClick={onApplyDirectTextEdits}
                  disabled={loading || !directEditMode}
                >
                  Apply Typed Changes
                </button>
                <button
                  type="button"
                  className="manual-reset-btn"
                  onClick={onDiscardDirectTextEdits}
                  disabled={loading}
                >
                  Discard Unapplied Typing
                </button>
              </div>
            </div>
          )}
          {!!state?.contextFiles.length && (
            <div className="context-list">
              <h3>Context Files</h3>
              {state.contextFiles.map((file) => (
                <div key={file.id} className="context-row">
                  <p>
                    {file.filename} ({file.charCount.toLocaleString()} chars)
                  </p>
                  <button
                    type="button"
                    className="danger-btn context-remove-btn"
                    onClick={() => onRemoveContextFile(file.id)}
                    disabled={loading}
                  >
                    Remove
                  </button>
                </div>
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
      )}
    </div>
  );
}

export default App;

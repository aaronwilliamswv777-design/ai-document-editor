import {
  CSSProperties,
  ChangeEvent,
  DragEvent,
  Fragment,
  SyntheticEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState
} from "react";
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
type SidebarSectionId =
  | "workflow"
  | "documents"
  | "prompt"
  | "changes"
  | "context"
  | "workspace"
  | "direct"
  | "ai"
  | "appearance";
type SidebarDropPreview = {
  targetId: SidebarSectionId;
  placement: "before" | "after";
};
type SidebarSectionOpenState = Record<SidebarSectionId, boolean>;
type ThemeSettings = {
  overallColor: string;
  mainUiColor: string;
  accentColor: string;
};
type PersistedPreferences = {
  provider?: Provider;
  model?: string;
  apiKeys?: Partial<ProviderKeyMap>;
};
const PREFERENCES_STORAGE_KEY = "doc-edit.preferences.v1";
const REMEMBER_SETTINGS_KEY = "doc-edit.remember-settings.v1";
const MODEL_CATALOG_STORAGE_KEY = "doc-edit.model-catalog.v1";
const EDIT_MODE_STORAGE_KEY = "doc-edit.ui.edit-mode.v1";
const TOP_MENU_STORAGE_KEY = "doc-edit.ui.top-menu.v1";
const SIDEBAR_ORDER_STORAGE_KEY = "doc-edit.ui.sidebar-order.v1";
const SIDEBAR_SECTION_OPEN_STORAGE_KEY = "doc-edit.ui.sidebar-section-open.v1";
const THEME_STORAGE_KEY = "doc-edit.ui.theme.v1";
const DEFAULT_THEME: ThemeSettings = {
  overallColor: "#090c12",
  mainUiColor: "#171c26",
  accentColor: "#2b9cff"
};
const DEFAULT_SIDEBAR_ORDER: SidebarSectionId[] = [
  "workflow",
  "documents",
  "prompt",
  "changes",
  "context",
  "workspace",
  "direct",
  "ai",
  "appearance"
];

type GrammarHighlight = {
  editId: string;
  mode: "grammar" | "custom";
  blockId: string;
  blockText: string;
  targetText: string;
  replacementText: string;
  tooltip: string;
  hintIndex: number;
  targetOffset?: number;
  wordChangeIndex?: number;
};

function formatDate(iso: string): string {
  return new Date(iso).toLocaleString();
}

function isProviderValue(value: unknown): value is Provider {
  return value === "anthropic" || value === "gemini" || value === "openrouter";
}

function isEditModeValue(value: unknown): value is EditMode {
  return value === "custom" || value === "grammar";
}

function isTopMenuValue(value: unknown): value is TopMenu {
  return value === "editor" || value === "settings";
}

function isSidebarSectionValue(value: unknown): value is SidebarSectionId {
  return (
    value === "workflow" ||
    value === "documents" ||
    value === "prompt" ||
    value === "changes" ||
    value === "context" ||
    value === "workspace" ||
    value === "direct" ||
    value === "ai" ||
    value === "appearance"
  );
}

function sanitizeSidebarOrder(value: unknown): SidebarSectionId[] {
  if (!Array.isArray(value)) {
    return [...DEFAULT_SIDEBAR_ORDER];
  }
  const unique = new Set<SidebarSectionId>();
  for (const item of value) {
    if (isSidebarSectionValue(item)) {
      unique.add(item);
    }
  }
  for (const fallback of DEFAULT_SIDEBAR_ORDER) {
    if (!unique.has(fallback)) {
      unique.add(fallback);
    }
  }
  return Array.from(unique);
}

function createDefaultSidebarSectionOpenState(): SidebarSectionOpenState {
  return DEFAULT_SIDEBAR_ORDER.reduce((accumulator, sectionId) => {
    accumulator[sectionId] = true;
    return accumulator;
  }, {} as SidebarSectionOpenState);
}

function sanitizeSidebarSectionOpenState(value: unknown): SidebarSectionOpenState {
  const next = createDefaultSidebarSectionOpenState();
  if (!value || typeof value !== "object") {
    return next;
  }
  const payload = value as Partial<Record<SidebarSectionId, unknown>>;
  for (const sectionId of DEFAULT_SIDEBAR_ORDER) {
    const raw = payload[sectionId];
    if (typeof raw === "boolean") {
      next[sectionId] = raw;
    }
  }
  return next;
}

function loadSidebarSectionOpenStateFromStorage(): SidebarSectionOpenState {
  if (typeof window === "undefined") {
    return createDefaultSidebarSectionOpenState();
  }
  try {
    const raw = window.localStorage.getItem(SIDEBAR_SECTION_OPEN_STORAGE_KEY);
    if (!raw) {
      return createDefaultSidebarSectionOpenState();
    }
    const parsed = JSON.parse(raw) as unknown;
    return sanitizeSidebarSectionOpenState(parsed);
  } catch {
    return createDefaultSidebarSectionOpenState();
  }
}

function isHexColor(value: string): boolean {
  return /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(value);
}

function normalizeThemeColor(input: unknown, fallback: string): string {
  if (typeof input !== "string") {
    return fallback;
  }
  const trimmed = input.trim();
  return isHexColor(trimmed) ? trimmed : fallback;
}

function sanitizeThemeSettings(value: unknown): ThemeSettings {
  if (!value || typeof value !== "object") {
    return { ...DEFAULT_THEME };
  }
  const payload = value as Partial<ThemeSettings>;
  return {
    overallColor: normalizeThemeColor(payload.overallColor, DEFAULT_THEME.overallColor),
    mainUiColor: normalizeThemeColor(payload.mainUiColor, DEFAULT_THEME.mainUiColor),
    accentColor: normalizeThemeColor(payload.accentColor, DEFAULT_THEME.accentColor)
  };
}

function hexToRgba(value: string, alpha: number): string {
  if (!isHexColor(value)) {
    return `rgba(43, 156, 255, ${alpha})`;
  }
  const trimmed = value.replace("#", "");
  const full =
    trimmed.length === 3
      ? `${trimmed[0]}${trimmed[0]}${trimmed[1]}${trimmed[1]}${trimmed[2]}${trimmed[2]}`
      : trimmed;
  const r = parseInt(full.slice(0, 2), 16);
  const g = parseInt(full.slice(2, 4), 16);
  const b = parseInt(full.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
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
  return extractAtomicTokenSpans(value).map((item) => item.value);
}

type TokenSpan = {
  value: string;
  start: number;
  end: number;
};

function extractAtomicTokenSpans(value: string): TokenSpan[] {
  const spans: TokenSpan[] = [];
  const tokenPattern = /[A-Za-z0-9']+|[.,;:!?]/g;
  let match: RegExpExecArray | null = tokenPattern.exec(value);
  while (match) {
    const token = (match[0] || "").trim();
    if (token) {
      spans.push({
        value: token,
        start: match.index,
        end: match.index + token.length
      });
    }
    match = tokenPattern.exec(value);
  }
  return spans;
}

function deriveWordChangesFromDiffHtml(diffHtml: string): Array<{
  from: string;
  to: string;
  start: number;
  end: number;
}> {
  if (!diffHtml.trim()) {
    return [];
  }

  const parser = new DOMParser();
  const parsed = parser.parseFromString(`<div>${diffHtml}</div>`, "text/html");
  const root = parsed.body.firstElementChild;
  if (!root) {
    return [];
  }

  const parts = Array.from(root.querySelectorAll("span")).map((node) => ({
    value: node.textContent || "",
    added: node.classList.contains("diff-added"),
    removed: node.classList.contains("diff-removed")
  }));

  const changes: Array<{
    from: string;
    to: string;
    start: number;
    end: number;
  }> = [];

  let cursor = 0;
  let index = 0;

  while (index < parts.length) {
    const part = parts[index];
    if (!part.added && !part.removed) {
      cursor += part.value.length;
      index += 1;
      continue;
    }

    const clusterStart = cursor;
    let removedSegment = "";
    let addedSegment = "";

    while (index < parts.length) {
      const next = parts[index];
      if (!next.added && !next.removed) {
        break;
      }
      if (next.removed) {
        removedSegment += next.value;
        cursor += next.value.length;
      } else if (next.added) {
        addedSegment += next.value;
      }
      index += 1;
    }

    const removedTokens = extractAtomicTokenSpans(removedSegment);
    if (removedTokens.length === 0) {
      continue;
    }

    const addedTokens = extractAtomicHighlightTargets(addedSegment);
    for (let tokenIndex = 0; tokenIndex < removedTokens.length; tokenIndex += 1) {
      const removedToken = removedTokens[tokenIndex];
      changes.push({
        from: removedToken.value,
        to: addedTokens[tokenIndex] || addedTokens[addedTokens.length - 1] || "(removed)",
        start: clusterStart + removedToken.start,
        end: clusterStart + removedToken.end
      });
    }
  }

  return changes;
}

function isWordToken(value: string): boolean {
  return /^[A-Za-z0-9']+$/.test(value);
}

function isWordChar(value: string): boolean {
  return /[A-Za-z0-9']/.test(value);
}

function findNeedleStartFrom(fullText: string, needle: string, fromIndex: number): number {
  const loweredText = fullText.toLowerCase();
  const loweredNeedle = needle.toLowerCase();
  let cursor = Math.max(0, Math.min(fromIndex, loweredText.length));

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

function findNeedleStart(fullText: string, needle: string, preferredStart?: number): number {
  if (typeof preferredStart === "number" && Number.isFinite(preferredStart)) {
    const hinted = findNeedleStartFrom(fullText, needle, Math.max(0, preferredStart - 24));
    if (hinted >= 0) {
      return hinted;
    }
  }
  return findNeedleStartFrom(fullText, needle, 0);
}

function applyHighlightToParagraph(paragraph: Element, highlight: GrammarHighlight): boolean {
  const needle = highlight.targetText.trim();
  if (!needle) {
    return false;
  }

  const textNodes = collectTextNodes(paragraph).filter((node) => (node.nodeValue || "").length > 0);
  if (textNodes.length === 0) {
    return false;
  }

  const fullText = textNodes.map((node) => node.nodeValue || "").join("");
  const start = findNeedleStart(fullText, needle, highlight.targetOffset);
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
  marker.setAttribute("data-grammar-tooltip", highlight.tooltip);
  marker.setAttribute("aria-label", highlight.tooltip);
  marker.setAttribute("title", highlight.tooltip);

  try {
    range.surroundContents(marker);
  } catch {
    const extracted = range.extractContents();
    marker.appendChild(extracted);
    range.insertNode(marker);
  }

  const anchor = document.createElement("span");
  anchor.className = "grammar-issue-anchor";
  marker.parentNode?.insertBefore(anchor, marker);
  anchor.appendChild(marker);

  const popover = document.createElement("span");
  popover.className = "grammar-inline-popover";
  popover.setAttribute("aria-hidden", "true");

  const issueType = document.createElement("span");
  issueType.className = "grammar-inline-popover-type";
  issueType.textContent = highlight.mode === "grammar" ? "Grammar change" : "Targeted change";
  popover.appendChild(issueType);

  const originalLabel = document.createElement("span");
  originalLabel.className = "grammar-inline-popover-label";
  originalLabel.textContent = "Original";
  popover.appendChild(originalLabel);

  const originalValue = document.createElement("span");
  originalValue.className = "grammar-inline-popover-body";
  originalValue.textContent = highlight.targetText;
  popover.appendChild(originalValue);

  const replacementLabel = document.createElement("span");
  replacementLabel.className = "grammar-inline-popover-label";
  replacementLabel.textContent = "Changed to";
  popover.appendChild(replacementLabel);

  const replacementValue = document.createElement("span");
  replacementValue.className = "grammar-inline-popover-body";
  replacementValue.textContent = highlight.replacementText;
  popover.appendChild(replacementValue);

  const actionRow = document.createElement("span");
  actionRow.className = "grammar-inline-actions";

  const acceptButton = document.createElement("button");
  acceptButton.type = "button";
  acceptButton.className = "inline-decision-btn inline-decision-accept";
  acceptButton.dataset.inlineDecision = "accept";
  acceptButton.dataset.editId = highlight.editId;
  if (typeof highlight.wordChangeIndex === "number") {
    acceptButton.dataset.wordChangeIndex = String(highlight.wordChangeIndex);
  }
  acceptButton.textContent = "Accept";
  actionRow.appendChild(acceptButton);

  const rejectButton = document.createElement("button");
  rejectButton.type = "button";
  rejectButton.className = "inline-decision-btn inline-decision-reject";
  rejectButton.dataset.inlineDecision = "reject";
  rejectButton.dataset.editId = highlight.editId;
  if (typeof highlight.wordChangeIndex === "number") {
    rejectButton.dataset.wordChangeIndex = String(highlight.wordChangeIndex);
  }
  rejectButton.textContent = "Reject";
  actionRow.appendChild(rejectButton);

  popover.appendChild(actionRow);
  anchor.appendChild(popover);

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
  const boundBlockElements = Array.from(
    container.querySelectorAll<HTMLElement>(".docx [data-block-id]")
  ).reduce((accumulator, element) => {
    const blockId = element.dataset.blockId;
    if (!blockId) {
      return accumulator;
    }
    const existing = accumulator.get(blockId);
    if (existing) {
      existing.push(element);
    } else {
      accumulator.set(blockId, [element]);
    }
    return accumulator;
  }, new Map<string, HTMLElement[]>());

  for (const highlight of highlights) {
    const boundCandidates = boundBlockElements.get(highlight.blockId) || [];
    if (boundCandidates.length > 0) {
      let appliedToBound = false;
      for (const candidate of boundCandidates) {
        if (applyHighlightToParagraph(candidate, highlight)) {
          appliedToBound = true;
          break;
        }
      }
      if (appliedToBound) {
        continue;
      }
    }

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

    applyHighlightToParagraph(paragraph, highlight);
  }
}

function App() {
  const uiRevision = "sidebar-modern-v18";
  const [sessionId, setSessionId] = useState<string>("");
  const [state, setState] = useState<SessionState | null>(null);
  const [instructionText, setInstructionText] = useState("");
  const [editMode, setEditMode] = useState<EditMode>("custom");
  const [topMenu, setTopMenu] = useState<TopMenu>("editor");
  const [sidebarOrder, setSidebarOrder] = useState<SidebarSectionId[]>([...DEFAULT_SIDEBAR_ORDER]);
  const [sidebarSectionOpen, setSidebarSectionOpen] = useState<SidebarSectionOpenState>(
    loadSidebarSectionOpenStateFromStorage
  );
  const [draggingSidebarSection, setDraggingSidebarSection] = useState<SidebarSectionId | null>(null);
  const [sidebarDropPreview, setSidebarDropPreview] = useState<SidebarDropPreview | null>(null);
  const [themeSettings, setThemeSettings] = useState<ThemeSettings>({ ...DEFAULT_THEME });
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

  const refresh = useCallback(async (targetSessionId: string): Promise<void> => {
    const next = await fetchState(targetSessionId);
    setState(next);
  }, []);

  useEffect(() => {
    try {
      const savedEditMode = window.localStorage.getItem(EDIT_MODE_STORAGE_KEY);
      if (isEditModeValue(savedEditMode)) {
        setEditMode(savedEditMode);
      }
      const savedTopMenu = window.localStorage.getItem(TOP_MENU_STORAGE_KEY);
      if (isTopMenuValue(savedTopMenu)) {
        setTopMenu(savedTopMenu);
      }

      const rawSidebarOrder = window.localStorage.getItem(SIDEBAR_ORDER_STORAGE_KEY);
      if (rawSidebarOrder) {
        const parsedOrder = JSON.parse(rawSidebarOrder) as unknown;
        setSidebarOrder(sanitizeSidebarOrder(parsedOrder));
      }

      const rawTheme = window.localStorage.getItem(THEME_STORAGE_KEY);
      if (rawTheme) {
        const parsedTheme = JSON.parse(rawTheme) as unknown;
        setThemeSettings(sanitizeThemeSettings(parsedTheme));
      }

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
  }, [refresh]);

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
      window.localStorage.setItem(EDIT_MODE_STORAGE_KEY, editMode);
      window.localStorage.setItem(TOP_MENU_STORAGE_KEY, topMenu);
      window.localStorage.setItem(SIDEBAR_ORDER_STORAGE_KEY, JSON.stringify(sidebarOrder));
      window.localStorage.setItem(
        SIDEBAR_SECTION_OPEN_STORAGE_KEY,
        JSON.stringify(sidebarSectionOpen)
      );
      window.localStorage.setItem(THEME_STORAGE_KEY, JSON.stringify(themeSettings));
    } catch {
      // Ignore browser storage failures.
    }
  }, [preferencesLoaded, editMode, topMenu, sidebarOrder, sidebarSectionOpen, themeSettings]);

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
      setError(`Add a ${provider} API key in AI Settings before running this action.`);
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

  const onDecide = useCallback(
    async (
      editId: string,
      decision: "accept" | "reject",
      wordChangeIndex?: number
    ): Promise<void> => {
      if (!sessionId) {
        return;
      }
      try {
        setLoading(true);
        setError("");
        await decideEdit(sessionId, editId, decision, wordChangeIndex);
        await refresh(sessionId);
        setStatus(
          typeof wordChangeIndex === "number"
            ? `Word change ${decision}ed.`
            : `Suggestion ${decision}ed.`
        );
      } catch (decisionError) {
        setError(decisionError instanceof Error ? decisionError.message : "Failed to apply decision.");
      } finally {
        setLoading(false);
      }
    },
    [sessionId, refresh]
  );

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

  function onSidebarSectionToggle(
    event: SyntheticEvent<HTMLDetailsElement>,
    sectionId: SidebarSectionId
  ): void {
    const nextOpen = event.currentTarget.open;
    setSidebarSectionOpen((previous) => {
      if (previous[sectionId] === nextOpen) {
        return previous;
      }
      const next = { ...previous, [sectionId]: nextOpen };
      try {
        window.localStorage.setItem(SIDEBAR_SECTION_OPEN_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // Ignore browser storage failures.
      }
      return next;
    });
  }

  function onSidebarSectionDragStart(
    event: DragEvent<HTMLElement>,
    sectionId: SidebarSectionId
  ): void {
    setDraggingSidebarSection(sectionId);
    setSidebarDropPreview(null);
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", sectionId);
  }

  function onSidebarSectionDragStartCapture(event: DragEvent<HTMLElement>): void {
    const target = event.target;
    if (target instanceof HTMLElement && target.closest(".drag-handle")) {
      return;
    }
    event.preventDefault();
  }

  function resolveSidebarDragSource(event: DragEvent<HTMLElement>): SidebarSectionId | null {
    const transferred = event.dataTransfer.getData("text/plain");
    const sourceId = isSidebarSectionValue(transferred) ? transferred : draggingSidebarSection;
    return sourceId || null;
  }

  function resolveSidebarDropPlacement(event: DragEvent<HTMLElement>): "before" | "after" {
    const bounds = event.currentTarget.getBoundingClientRect();
    const midpoint = bounds.top + bounds.height / 2;
    return event.clientY < midpoint ? "before" : "after";
  }

  function computeSidebarReorder(
    order: SidebarSectionId[],
    sourceId: SidebarSectionId,
    targetId: SidebarSectionId,
    placement: "before" | "after"
  ): SidebarSectionId[] | null {
    if (sourceId === targetId) {
      return null;
    }

    const sourceIndex = order.indexOf(sourceId);
    const targetIndex = order.indexOf(targetId);
    if (sourceIndex < 0 || targetIndex < 0) {
      return null;
    }

    const next = [...order];
    next.splice(sourceIndex, 1);

    const normalizedTargetIndex = next.indexOf(targetId);
    if (normalizedTargetIndex < 0) {
      return null;
    }

    const insertionIndex = placement === "after" ? normalizedTargetIndex + 1 : normalizedTargetIndex;
    next.splice(insertionIndex, 0, sourceId);
    const changed = next.some((item, index) => item !== order[index]);
    return changed ? next : null;
  }

  function moveSidebarSection(
    sourceId: SidebarSectionId,
    targetId: SidebarSectionId,
    placement: "before" | "after"
  ): void {
    setSidebarOrder((previous) => {
      const next = computeSidebarReorder(previous, sourceId, targetId, placement);
      return next || previous;
    });
  }

  function onSidebarSectionDragOver(
    event: DragEvent<HTMLElement>,
    targetId: SidebarSectionId
  ): void {
    const sourceId = resolveSidebarDragSource(event);
    if (!sourceId || sourceId === targetId) {
      setSidebarDropPreview(null);
      return;
    }
    const placement = resolveSidebarDropPlacement(event);
    if (!computeSidebarReorder(sidebarOrder, sourceId, targetId, placement)) {
      setSidebarDropPreview(null);
      return;
    }
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
    setSidebarDropPreview((previous) =>
      previous?.targetId === targetId && previous.placement === placement
        ? previous
        : { targetId, placement }
    );
  }

  function onSidebarSectionDrop(event: DragEvent<HTMLElement>, targetId: SidebarSectionId): void {
    const sourceId = resolveSidebarDragSource(event);
    if (!sourceId || sourceId === targetId) {
      setDraggingSidebarSection(null);
      setSidebarDropPreview(null);
      return;
    }
    event.preventDefault();
    const preview = sidebarDropPreview?.targetId === targetId ? sidebarDropPreview : null;
    const placement = preview ? preview.placement : resolveSidebarDropPlacement(event);
    moveSidebarSection(sourceId, targetId, placement);
    setDraggingSidebarSection(null);
    setSidebarDropPreview(null);
  }

  function onSidebarGhostDragOver(
    event: DragEvent<HTMLElement>,
    targetId: SidebarSectionId,
    placement: "before" | "after"
  ): void {
    const sourceId = resolveSidebarDragSource(event);
    if (!sourceId || sourceId === targetId) {
      setSidebarDropPreview(null);
      return;
    }
    if (!computeSidebarReorder(sidebarOrder, sourceId, targetId, placement)) {
      setSidebarDropPreview(null);
      return;
    }
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
    setSidebarDropPreview((previous) =>
      previous?.targetId === targetId && previous.placement === placement
        ? previous
        : { targetId, placement }
    );
  }

  function onSidebarGhostDrop(
    event: DragEvent<HTMLElement>,
    targetId: SidebarSectionId,
    placement: "before" | "after"
  ): void {
    const sourceId = resolveSidebarDragSource(event);
    if (!sourceId || sourceId === targetId) {
      setDraggingSidebarSection(null);
      setSidebarDropPreview(null);
      return;
    }
    event.preventDefault();
    moveSidebarSection(sourceId, targetId, placement);
    setDraggingSidebarSection(null);
    setSidebarDropPreview(null);
  }

  function onSidebarSectionDragEnd(): void {
    setDraggingSidebarSection(null);
    setSidebarDropPreview(null);
  }

  function onResetThemeDefaults(): void {
    setThemeSettings({ ...DEFAULT_THEME });
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

  const pendingEdits = useMemo(() => {
    if (!state) {
      return [] as Array<{
        batch: ProposalBatch;
        edit: ProposalBatch["edits"][number];
      }>;
    }
    return allBatches.flatMap((batch) =>
      batch.edits
        .filter((edit) => edit.status === "pending")
        .map((edit) => ({
          batch,
          edit
        }))
    );
  }, [state, allBatches]);

  const appThemeStyle = useMemo(
    () =>
      ({
        "--app-bg": themeSettings.overallColor,
        "--main-ui": themeSettings.mainUiColor,
        "--accent": themeSettings.accentColor,
        "--accent-ghost": hexToRgba(themeSettings.accentColor, 0.16),
        "--accent-glow": hexToRgba(themeSettings.accentColor, 0.45)
      }) as CSSProperties,
    [themeSettings]
  );

  const grammarHighlights = useMemo(() => {
    if (!state) {
      return [] as GrammarHighlight[];
    }

    const blockMap = new Map(state.workingBlocks.map((block) => [block.id, block]));
    const blockIndexById = new Map(state.workingBlocks.map((block, index) => [block.id, index]));
    const highlights: GrammarHighlight[] = [];

    for (const batch of state.proposalHistory) {
      for (const edit of batch.edits) {
        if (edit.status !== "pending") {
          continue;
        }
        const block = blockMap.get(edit.blockId);
        if (!block) {
          continue;
        }
        const hintIndex = blockIndexById.get(edit.blockId) ?? 0;

        const derivedWordChanges =
          Array.isArray(edit.wordChanges) && edit.wordChanges.length > 0
            ? edit.wordChanges
            : deriveWordChangesFromDiffHtml(edit.diffHtml);
        const wordChangeStatuses =
          Array.isArray(edit.wordChangeStatuses) &&
          edit.wordChangeStatuses.length === derivedWordChanges.length
            ? edit.wordChangeStatuses
            : derivedWordChanges.map(() => "pending" as const);

        if (derivedWordChanges.length > 0) {
          for (let wordChangeIndex = 0; wordChangeIndex < derivedWordChanges.length; wordChangeIndex += 1) {
            if (wordChangeStatuses[wordChangeIndex] !== "pending") {
              continue;
            }
            const wordChange = derivedWordChanges[wordChangeIndex];
            const fromToken = typeof wordChange.from === "string" ? wordChange.from.trim() : "";
            const toTokenRaw = typeof wordChange.to === "string" ? wordChange.to.trim() : "";
            if (!fromToken) {
              continue;
            }
            const toToken = toTokenRaw || "(removed)";
            highlights.push({
              editId: edit.id,
              mode: batch.mode,
              blockId: edit.blockId,
              blockText: block.text,
              targetText: fromToken,
              replacementText: toToken,
              tooltip: `Original: ${fromToken} | Changed to: ${toToken}`,
              hintIndex,
              targetOffset:
                typeof wordChange.start === "number" && Number.isFinite(wordChange.start)
                  ? wordChange.start
                  : undefined,
              wordChangeIndex
            });
          }
          continue;
        }

        const rawTargetTexts = edit.highlightTexts?.length
          ? edit.highlightTexts
          : edit.highlightText
            ? [edit.highlightText]
            : [];

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
          highlights.push({
            editId: edit.id,
            mode: batch.mode,
            blockId: edit.blockId,
            blockText: block.text,
            targetText,
            replacementText: "(changed)",
            tooltip: `Original: ${targetText} | Changed to: (changed)`,
            hintIndex
          });
        }
      }
    }

    return highlights.sort((left, right) => {
      if (left.hintIndex !== right.hintIndex) {
        return right.hintIndex - left.hintIndex;
      }
      const leftOffset = left.targetOffset ?? Number.NEGATIVE_INFINITY;
      const rightOffset = right.targetOffset ?? Number.NEGATIVE_INFINITY;
      if (leftOffset !== rightOffset) {
        return rightOffset - leftOffset;
      }
      return right.targetText.localeCompare(left.targetText);
    });
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
        bindEditablePreviewBlocks(previewContainerRef.current, state.workingBlocks, directEditMode);
        if (!directEditMode) {
          applyGrammarHighlights(previewContainerRef.current, grammarHighlights);
        }
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

  useEffect(() => {
    if (topMenu !== "editor" || !previewContainerRef.current) {
      return;
    }

    const container = previewContainerRef.current;
    const openClassName = "grammar-inline-popover-open";
    const closeOpenInlinePopovers = (keepOpen?: HTMLElement): void => {
      const openAnchors = container.querySelectorAll<HTMLElement>(`.grammar-issue-anchor.${openClassName}`);
      for (const anchor of openAnchors) {
        if (keepOpen && anchor === keepOpen) {
          continue;
        }
        anchor.classList.remove(openClassName);
      }
    };
    const onInlineActionClick = (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const actionButton = target.closest<HTMLButtonElement>(".inline-decision-btn");
      if (actionButton) {
        event.preventDefault();
        event.stopPropagation();
        closeOpenInlinePopovers();

        const editId = actionButton.dataset.editId;
        const decision = actionButton.dataset.inlineDecision;
        const rawWordChangeIndex = actionButton.dataset.wordChangeIndex;
        const parsedWordChangeIndex =
          typeof rawWordChangeIndex === "string" && rawWordChangeIndex.length > 0
            ? Number.parseInt(rawWordChangeIndex, 10)
            : Number.NaN;
        const wordChangeIndex = Number.isFinite(parsedWordChangeIndex)
          ? parsedWordChangeIndex
          : undefined;
        if (!editId || loading || (decision !== "accept" && decision !== "reject")) {
          return;
        }
        void onDecide(editId, decision, wordChangeIndex);
        return;
      }

      const clickedHighlight = target.closest<HTMLElement>(".grammar-issue-highlight");
      const clickedAnchor = target.closest<HTMLElement>(".grammar-issue-anchor");
      if (clickedHighlight && clickedAnchor) {
        event.preventDefault();
        event.stopPropagation();
        const shouldOpen = !clickedAnchor.classList.contains(openClassName);
        closeOpenInlinePopovers(clickedAnchor);
        clickedAnchor.classList.toggle(openClassName, shouldOpen);
        return;
      }

      if (target.closest(".grammar-inline-popover")) {
        return;
      }

      closeOpenInlinePopovers();
    };

    container.addEventListener("click", onInlineActionClick);
    return () => {
      container.removeEventListener("click", onInlineActionClick);
    };
  }, [topMenu, onDecide, loading]);

  const editorSections = new Set<SidebarSectionId>([
    "workflow",
    "documents",
    "prompt",
    "changes",
    "context",
    "workspace",
    "direct",
    "appearance"
  ]);
  const settingsSections = new Set<SidebarSectionId>(["workflow", "ai", "appearance"]);
  const visibleSidebarSections = sidebarOrder.filter((id) =>
    topMenu === "editor" ? editorSections.has(id) : settingsSections.has(id)
  );

  return (
    <div className="app app-shell" style={appThemeStyle}>
      <aside className="sidebar" onDragStartCapture={onSidebarSectionDragStartCapture}>
        <header className="sidebar-header">
          <h1>Doc Editing Application</h1>
          <p className="subtle">{status}</p>
          <p className="subtle">
            Pending suggestions: {pendingCount} | UI rev: {uiRevision}
          </p>
          <p className="subtle">Drag section cards to reorder. Layout saves automatically.</p>
        </header>

        {error && <p className="error">{error}</p>}

        {visibleSidebarSections.map((sectionId) => {
          const showDropBefore =
            sidebarDropPreview?.targetId === sectionId && sidebarDropPreview.placement === "before";
          const showDropAfter =
            sidebarDropPreview?.targetId === sectionId && sidebarDropPreview.placement === "after";

          return (
            <Fragment key={sectionId}>
              {showDropBefore && (
                <div
                  className="sidebar-drop-ghost"
                  onDragOver={(event) => onSidebarGhostDragOver(event, sectionId, "before")}
                  onDrop={(event) => onSidebarGhostDrop(event, sectionId, "before")}
                >
                  Drop Here
                </div>
              )}
              <details
                className={`sidebar-section ${draggingSidebarSection === sectionId ? "sidebar-section-dragging" : ""}`}
                open={sidebarSectionOpen[sectionId]}
                onToggle={(event) => onSidebarSectionToggle(event, sectionId)}
                onDragOver={(event) => onSidebarSectionDragOver(event, sectionId)}
                onDrop={(event) => onSidebarSectionDrop(event, sectionId)}
              >
                <summary className="sidebar-section-summary">
                  <span>
                    {sectionId === "workflow" && "Workflow"}
                    {sectionId === "documents" && "Documents"}
                    {sectionId === "prompt" && "Prompt"}
                    {sectionId === "changes" && "Suggestions"}
                    {sectionId === "context" && "Context Files"}
                    {sectionId === "workspace" && "Workspace Controls"}
                    {sectionId === "direct" && "Manual Typing"}
                    {sectionId === "ai" && "AI Settings"}
                    {sectionId === "appearance" && "Appearance"}
                  </span>
                  <button
                    type="button"
                    className="drag-hint drag-handle"
                    draggable
                    onDragStart={(event) => onSidebarSectionDragStart(event, sectionId)}
                    onDragEnd={onSidebarSectionDragEnd}
                    onMouseDown={(event) => event.stopPropagation()}
                    onClick={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                    }}
                    aria-label="Drag section"
                  >
                    Drag
                  </button>
                </summary>
                <div className="sidebar-section-body">
              {sectionId === "workflow" && (
                <div className="stack">
                  <label>
                    Workspace
                    <select
                      value={topMenu}
                      onChange={(event) => setTopMenu(event.target.value as TopMenu)}
                      disabled={loading || loadingModels}
                    >
                      <option value="editor">Document Editor</option>
                      <option value="settings">AI Settings</option>
                    </select>
                  </label>
                  <label>
                    Edit Mode
                    <select
                      value={editMode}
                      onChange={(event) => setEditMode(event.target.value as EditMode)}
                      disabled={loading}
                    >
                      <option value="custom">Targeted Edit</option>
                      <option value="grammar">Grammar & Punctuation</option>
                    </select>
                  </label>
                  <p className="subtle">
                    Provider: {provider} | Model: {model.trim() || "Provider default"}
                  </p>
                </div>
              )}

              {sectionId === "documents" && (
                <div className="stack">
                  <label className="file-control">
                    Load main `.docx` file
                    <input
                      type="file"
                      accept=".docx"
                      onChange={onSourceFileChange}
                      disabled={loading}
                    />
                  </label>
                  <p className="subtle">
                    Active document: {state?.sourceFilename || "No source document loaded"}
                  </p>
                </div>
              )}

              {sectionId === "prompt" && (
                <div className="stack">
                  <label className="prompt-field">
                    {editMode === "custom"
                      ? "Tell AI what to change"
                      : "Grammar mode instructions (optional)"}
                    <textarea
                      value={instructionText}
                      onChange={(event) => setInstructionText(event.target.value)}
                      placeholder={
                        editMode === "custom"
                          ? "Example: tighten wording, improve clarity, keep executive tone."
                          : "Optional: focus on punctuation consistency and formal business grammar."
                      }
                      rows={5}
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
                    {editMode === "custom"
                      ? "Generate Suggested Edits"
                      : "Find Grammar & Punctuation Issues"}
                  </button>
                </div>
              )}

              {sectionId === "changes" && (
                <div className="stack">
                  <button
                    type="button"
                    onClick={onAcceptAllPending}
                    disabled={loading || pendingCount === 0}
                  >
                    Accept All Suggestions
                  </button>
                  {!pendingEdits.length && (
                    <p className="empty">No pending suggestions. Run analysis to generate edits.</p>
                  )}
                  {pendingEdits.map(({ batch, edit }) => (
                    <details key={edit.id} className="change-row">
                      <summary>
                        <span>
                          {batch.mode === "grammar" ? "Grammar" : "Targeted"}:{" "}
                          {(edit.highlightTexts?.[0] || edit.highlightText || edit.originalText).slice(0, 72)}
                        </span>
                        <span className="badge">{edit.status}</span>
                      </summary>
                      <div className="change-row-body">
                        <p className="subtle">
                          {formatDate(batch.createdAt)} | {batch.provider} | {batch.model}
                        </p>
                        <div
                          className="diff-content"
                          dangerouslySetInnerHTML={{
                            __html: edit.diffHtml
                          }}
                        />
                        <p className="rationale">{edit.rationale}</p>
                        <div className="actions">
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
                    </details>
                  ))}
                </div>
              )}

              {sectionId === "context" && (
                <div className="stack">
                  <label className="file-control">
                    Add optional context files
                    <input
                      type="file"
                      accept=".docx,.txt,.md"
                      multiple
                      onChange={onContextFileChange}
                      disabled={loading}
                    />
                  </label>
                  {!state?.contextFiles.length && (
                    <p className="subtle">No context files uploaded for this session.</p>
                  )}
                  {!!state?.contextFiles.length && (
                    <div className="context-list">
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
                </div>
              )}

              {sectionId === "workspace" && (
                <div className="stack">
                  <button
                    type="button"
                    className="session-save-btn"
                    onClick={onSaveWorkspaceForReturn}
                    disabled={loading || !state?.workingBlocks.length}
                  >
                    Save Workspace for Next Launch
                  </button>
                  <button
                    type="button"
                    className="session-remove-saved-btn"
                    onClick={onRemoveSavedWorkspaceForReturn}
                    disabled={loading}
                  >
                    Remove Saved Workspace Snapshot
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
              )}

              {sectionId === "direct" && (
                <div className="stack">
                  {!state?.workingBlocks.length && (
                    <p className="subtle">Load a document to enable direct typing mode.</p>
                  )}
                  {!!state?.workingBlocks.length && (
                    <>
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
                        Apply Manual Edits to Working Document
                      </button>
                      <button
                        type="button"
                        className="manual-reset-btn"
                        onClick={onDiscardDirectTextEdits}
                        disabled={loading}
                      >
                        Discard Unapplied Manual Typing
                      </button>
                    </>
                  )}
                </div>
              )}

              {sectionId === "ai" && (
                <div className="stack">
                  <label>
                    AI Provider
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
                  <button
                    type="button"
                    className="api-key-load-btn"
                    onClick={onLoadModels}
                    disabled={loading || loadingModels}
                  >
                    {loadingModels ? "Loading Available Models..." : `Load Models for ${provider}`}
                  </button>
                  <label>
                    Available models
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
                          : "Select a model from this list"}
                      </option>
                      {filteredProviderModels.length === 0 && providerModels[provider].length > 0 && (
                        <option value="__no_match__" disabled>
                          No models match typed text
                        </option>
                      )}
                      {filteredProviderModels.map((item) => (
                        <option key={item.id} value={item.id}>
                          {item.label ? `${item.label} (${item.id})` : item.id}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Model override text
                    <input
                      type="text"
                      className="model-search-input"
                      value={model}
                      onChange={(event) => setModel(event.target.value)}
                      placeholder="Type model id to filter/select"
                      autoComplete="off"
                      spellCheck={false}
                      disabled={loading || loadingModels}
                    />
                  </label>
                  <p className="subtle">
                    Showing {filteredProviderModels.length} of {providerModels[provider].length} models
                  </p>

                  <label className="remember-toggle">
                    <input
                      type="checkbox"
                      checked={rememberPreferences}
                      onChange={(event) => setRememberPreferences(event.target.checked)}
                      disabled={loading || loadingModels}
                    />
                    Remember API keys and model on this device
                  </label>

                  <button
                    type="button"
                    className={`api-key-toggle ${showApiKeyMenu ? "api-key-toggle-active" : ""}`}
                    onClick={() => setShowApiKeyMenu((prev) => !prev)}
                    disabled={loading || loadingModels}
                  >
                    {showApiKeyMenu ? "Hide API Key Inputs" : "Show API Key Inputs"}
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
                    </div>
                  )}
                  <button
                    type="button"
                    className="api-key-clear-btn"
                    onClick={onClearSavedPreferences}
                    disabled={loading || loadingModels}
                  >
                    Clear Saved AI Credentials and Model
                  </button>
                </div>
              )}

              {sectionId === "appearance" && (
                <div className="stack">
                  <label>
                    Overall app color
                    <input
                      type="color"
                      value={themeSettings.overallColor}
                      onChange={(event) =>
                        setThemeSettings((prev) => ({
                          ...prev,
                          overallColor: event.target.value
                        }))
                      }
                      disabled={loading || loadingModels}
                    />
                  </label>
                  <label>
                    Main UI panel color
                    <input
                      type="color"
                      value={themeSettings.mainUiColor}
                      onChange={(event) =>
                        setThemeSettings((prev) => ({
                          ...prev,
                          mainUiColor: event.target.value
                        }))
                      }
                      disabled={loading || loadingModels}
                    />
                  </label>
                  <label>
                    Button accent color
                    <input
                      type="color"
                      value={themeSettings.accentColor}
                      onChange={(event) =>
                        setThemeSettings((prev) => ({
                          ...prev,
                          accentColor: event.target.value
                        }))
                      }
                      disabled={loading || loadingModels}
                    />
                  </label>
                  <button type="button" onClick={onResetThemeDefaults} disabled={loading || loadingModels}>
                    Restore Default Theme
                  </button>
                </div>
              )}
                </div>
              </details>
              {showDropAfter && (
                <div
                  className="sidebar-drop-ghost"
                  onDragOver={(event) => onSidebarGhostDragOver(event, sectionId, "after")}
                  onDrop={(event) => onSidebarGhostDrop(event, sectionId, "after")}
                >
                  Drop Here
                </div>
              )}
            </Fragment>
          );
        })}
      </aside>

      <main className="main-stage">
        {topMenu === "editor" ? (
          <section className="document-stage">
            <header className="document-stage-header">
              <h2>Current Working Document</h2>
              <p className="subtle">
                Click red highlights to review why each word changes. Accept or reject directly in
                the document or from the sidebar suggestions list.
              </p>
            </header>
            {!state?.workingBlocks.length && (
              <p className="empty">Load a `.docx` file to begin reviewing and editing.</p>
            )}
            {previewLoading && <p className="subtle">Rendering formatted Word preview...</p>}
            {previewError && <p className="error">{previewError}</p>}
            <div className="doc-view doc-view-formatted">
              <div ref={previewContainerRef} className="docx-host" />
            </div>
          </section>
        ) : (
          <section className="settings-stage">
            <h2>AI Settings</h2>
            <p className="subtle">
              Use the sidebar cards to configure model provider, API keys, edit mode, and color
              preferences.
            </p>
            <div className="settings-summary-grid">
              <article className="settings-summary-card">
                <h3>Active Workspace</h3>
                <p className="subtle">View: AI Settings</p>
                <p className="subtle">
                  Edit mode: {editMode === "grammar" ? "Grammar & Punctuation" : "Targeted Edit"}
                </p>
              </article>
              <article className="settings-summary-card">
                <h3>Model Selection</h3>
                <p className="subtle">Provider: {provider}</p>
                <p className="subtle">Model: {model.trim() || "Provider default"}</p>
                <p className="subtle">
                  Saved models for provider: {providerModels[provider].length.toLocaleString()}
                </p>
              </article>
              <article className="settings-summary-card">
                <h3>Credential State</h3>
                <p className="subtle">Provider key present: {apiKeys[provider].trim() ? "Yes" : "No"}</p>
                <p className="subtle">
                  Remember setting: {rememberPreferences ? "On (stored locally)" : "Off"}
                </p>
              </article>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

export default App;


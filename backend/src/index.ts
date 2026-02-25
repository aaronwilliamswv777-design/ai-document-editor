import "dotenv/config";
import cors from "cors";
import express from "express";
import multer from "multer";
import { v4 as uuidv4 } from "uuid";
import { z } from "zod";
import { diffWords } from "diff";
import { createSession, deleteSession, getSession, updateSession } from "./sessionStore.js";
import {
  applyEditsToDocxBuffer,
  buildDiffHtml,
  extractTextFromGenericContext,
  parseDocxToBlocks
} from "./services/documentService.js";
import { generateEdits, listProviderModels } from "./services/aiService.js";
import {
  clearSavedWorkspaceSnapshot,
  loadSavedWorkspaceSnapshot,
  saveWorkspaceSnapshot
} from "./services/workspaceSaveService.js";
import { EditStatus, ProposalBatch, WordChange } from "./types.js";

const app = express();
const port = Number(process.env.PORT || 8080);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }
});

app.use(cors());
app.use(express.json({ limit: "1mb" }));

const proposeSchema = z.object({
  prompt: z.string().min(3),
  provider: z.enum(["anthropic", "gemini", "openrouter"]).optional(),
  model: z.string().optional(),
  apiKey: z.string().min(10)
});

const grammarAnalyzeSchema = z.object({
  customInstructions: z.string().max(4000).optional(),
  provider: z.enum(["anthropic", "gemini", "openrouter"]).optional(),
  model: z.string().optional(),
  apiKey: z.string().min(10)
});

const decisionSchema = z.object({
  decision: z.enum(["accept", "reject"]),
  wordChangeIndex: z.number().int().min(0).optional()
});

const manualEditSchema = z.object({
  edits: z
    .array(
      z.object({
        blockId: z.string().min(1),
        text: z.string().max(30_000)
      })
    )
    .min(1)
    .max(5000)
});

const listModelsSchema = z.object({
  provider: z.enum(["anthropic", "gemini", "openrouter"]),
  apiKey: z.string().min(10)
});

function cloneBlocks<T extends { id: string; text: string }>(blocks: T[]): T[] {
  return blocks.map((block) => ({ ...block }));
}

function cloneBuffer(buffer?: Buffer): Buffer | undefined {
  if (!buffer) {
    return undefined;
  }
  return Buffer.from(buffer);
}

function readRouteParam(value: string | string[] | undefined): string {
  if (!value) {
    return "";
  }
  if (Array.isArray(value)) {
    return value[0] || "";
  }
  return value;
}

function requireDocx(filename: string): boolean {
  return filename.toLowerCase().endsWith(".docx");
}

function trimContext(text: string): string {
  if (text.length <= 20_000) {
    return text;
  }
  return `${text.slice(0, 20_000)}\n\n[Truncated for token safety]`;
}

function buildGrammarPrompt(customInstructions: string | undefined, allowSentenceLevelChanges: boolean): string {
  const base = [
    "Analyze this document for grammar and punctuation issues only.",
    "Do not make stylistic rewrites or alter meaning."
  ].join("\n");

  const strictConstraints = allowSentenceLevelChanges
    ? "Apply only the sentence-level edits explicitly requested by the user."
    : [
        "Never delete or rewrite full sentences.",
        "Keep sentence structure and meaning intact.",
        "Only make local word-level or punctuation-level fixes."
      ].join("\n");

  if (!customInstructions?.trim()) {
    return `${base}\n${strictConstraints}`;
  }

  return `${base}\n${strictConstraints}\n\nAdditional instructions:\n${customInstructions.trim()}`;
}

function countSentenceMarkers(text: string): number {
  const matches = text.match(/[.!?](?=\s|$)/g);
  return matches ? matches.length : 0;
}

function countWordTokens(text: string): number {
  const matches = text.match(/[A-Za-z0-9']+/g);
  return matches ? matches.length : 0;
}

function allowsSentenceLevelChanges(customInstructions?: string): boolean {
  if (!customInstructions?.trim()) {
    return false;
  }
  return /\b(delete|remove|rewrite|reword|shorten|condense|summari[sz]e|merge|split|replace sentence|drop sentence)\b/i.test(
    customInstructions
  );
}

function isGrammarSafeEdit(
  originalText: string,
  proposedText: string,
  allowSentenceLevelChanges: boolean
): boolean {
  const original = originalText.trim();
  const proposed = proposedText.trim();

  if (!original || !proposed || original === proposed) {
    return false;
  }
  if (allowSentenceLevelChanges) {
    return true;
  }

  const originalSentenceCount = countSentenceMarkers(original);
  const proposedSentenceCount = countSentenceMarkers(proposed);
  if (proposedSentenceCount < originalSentenceCount) {
    return false;
  }

  const originalWordCount = countWordTokens(original);
  const proposedWordCount = countWordTokens(proposed);
  if (originalWordCount >= 10 && proposedWordCount < Math.floor(originalWordCount * 0.75)) {
    return false;
  }

  const parts = diffWords(original, proposed);
  let removedWordCount = 0;
  let unchangedWordCount = 0;
  let maxRemovedRun = 0;
  let currentRemovedRun = 0;

  for (const part of parts) {
    const tokens = countWordTokens(part.value);
    if (part.removed) {
      removedWordCount += tokens;
      currentRemovedRun += tokens;
      if (currentRemovedRun > maxRemovedRun) {
        maxRemovedRun = currentRemovedRun;
      }
    } else if (part.added) {
      currentRemovedRun = 0;
    } else {
      unchangedWordCount += tokens;
      currentRemovedRun = 0;
    }
  }

  if (originalWordCount >= 8 && unchangedWordCount <= 1) {
    return false;
  }
  if (removedWordCount > Math.max(6, Math.floor(originalWordCount * 0.35))) {
    return false;
  }
  if (maxRemovedRun >= 8) {
    return false;
  }

  return true;
}

function extractAtomicTokens(value: string): string[] {
  const matches = value.match(/[A-Za-z0-9']+|[.,;:!?]/g);
  if (!matches) {
    return [];
  }
  return matches.map((token) => token.trim()).filter(Boolean);
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

function deriveWordChanges(originalText: string, proposedText: string): WordChange[] {
  const changes: WordChange[] = [];
  const parts = diffWords(originalText, proposedText);
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
      const nextPart = parts[index];
      if (!nextPart.added && !nextPart.removed) {
        break;
      }
      if (nextPart.removed) {
        removedSegment += nextPart.value;
        cursor += nextPart.value.length;
      } else if (nextPart.added) {
        addedSegment += nextPart.value;
      }
      index += 1;
    }

    const removedTokens = extractAtomicTokenSpans(removedSegment);
    if (removedTokens.length === 0) {
      continue;
    }

    const addedTokens = extractAtomicTokens(addedSegment);
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

function deriveHighlightTexts(wordChanges: WordChange[]): string[] {
  return wordChanges.map((change) => change.from).filter((value) => value.length > 0);
}

function isWordToken(value: string): boolean {
  return /^[A-Za-z0-9']+$/.test(value);
}

function isWordChar(value: string): boolean {
  return /[A-Za-z0-9']/.test(value);
}

function hasTokenBoundaries(text: string, start: number, token: string): boolean {
  if (!isWordToken(token)) {
    return true;
  }
  const end = start + token.length;
  const prev = start > 0 ? text[start - 1] : "";
  const next = end < text.length ? text[end] : "";
  return (!prev || !isWordChar(prev)) && (!next || !isWordChar(next));
}

function collectTokenStarts(text: string, token: string, ignoreCase: boolean): number[] {
  if (!token.length) {
    return [];
  }
  const haystack = ignoreCase ? text.toLowerCase() : text;
  const needle = ignoreCase ? token.toLowerCase() : token;
  const starts: number[] = [];
  let cursor = 0;

  while (cursor <= haystack.length - needle.length) {
    const index = haystack.indexOf(needle, cursor);
    if (index < 0) {
      break;
    }
    starts.push(index);
    cursor = index + 1;
  }

  return starts;
}

function findBestTokenStart(text: string, token: string, expectedStart: number): number {
  if (!token.length) {
    return -1;
  }

  const tryDirect = (ignoreCase: boolean): number => {
    if (expectedStart < 0 || expectedStart + token.length > text.length) {
      return -1;
    }
    const slice = text.slice(expectedStart, expectedStart + token.length);
    const matched = ignoreCase ? slice.toLowerCase() === token.toLowerCase() : slice === token;
    if (!matched || !hasTokenBoundaries(text, expectedStart, token)) {
      return -1;
    }
    return expectedStart;
  };

  const exactDirect = tryDirect(false);
  if (exactDirect >= 0) {
    return exactDirect;
  }
  const caseInsensitiveDirect = tryDirect(true);
  if (caseInsensitiveDirect >= 0) {
    return caseInsensitiveDirect;
  }

  const boundaryCandidates = (
    candidates: number[]
  ): number[] => candidates.filter((start) => hasTokenBoundaries(text, start, token));

  const exactCandidates = boundaryCandidates(collectTokenStarts(text, token, false));
  const candidates =
    exactCandidates.length > 0
      ? exactCandidates
      : boundaryCandidates(collectTokenStarts(text, token, true));
  if (candidates.length === 0) {
    return -1;
  }

  let best = candidates[0];
  let bestDistance = Math.abs(best - expectedStart);
  for (let index = 1; index < candidates.length; index += 1) {
    const start = candidates[index];
    const distance = Math.abs(start - expectedStart);
    if (distance < bestDistance || (distance === bestDistance && start < best)) {
      best = start;
      bestDistance = distance;
    }
  }

  return best;
}

function normalizeWordChangeStatuses(wordChanges: WordChange[], existing?: EditStatus[]): EditStatus[] {
  if (
    Array.isArray(existing) &&
    existing.length === wordChanges.length &&
    existing.every((item) => item === "pending" || item === "accepted" || item === "rejected")
  ) {
    return [...existing];
  }
  return wordChanges.map(() => "pending");
}

function applySingleWordChange(
  currentText: string,
  change: WordChange,
  allWordChanges: WordChange[],
  statuses: EditStatus[]
): string | null {
  const replacement = change.to === "(removed)" ? "" : change.to;
  let delta = 0;
  for (let index = 0; index < allWordChanges.length; index += 1) {
    const priorChange = allWordChanges[index];
    if (statuses[index] !== "accepted" || priorChange.start >= change.start) {
      continue;
    }
    const priorReplacement = priorChange.to === "(removed)" ? "" : priorChange.to;
    delta += priorReplacement.length - priorChange.from.length;
  }

  const expectedStart = change.start + delta;
  const start = findBestTokenStart(currentText, change.from, expectedStart);
  if (start < 0) {
    return null;
  }

  const end = start + change.from.length;
  return `${currentText.slice(0, start)}${replacement}${currentText.slice(end)}`;
}

function resolveStatusFromWordChanges(statuses: EditStatus[]): EditStatus {
  if (statuses.some((status) => status === "pending")) {
    return "pending";
  }
  return statuses.some((status) => status === "accepted") ? "accepted" : "rejected";
}

async function acceptPendingEditsForSession(session: NonNullable<ReturnType<typeof getSession>>): Promise<number> {
  const pendingEdits = session.proposalHistory.flatMap((batch) =>
    batch.edits.filter((edit) => edit.status === "pending")
  );

  if (pendingEdits.length === 0) {
    return 0;
  }

  if (!session.workingDocxBuffer) {
    throw new Error("Working document buffer is missing.");
  }

  const updates: Array<{ blockId: string; text: string }> = [];
  for (const edit of pendingEdits) {
    edit.status = "accepted";
    if (Array.isArray(edit.wordChanges) && edit.wordChanges.length > 0) {
      edit.wordChangeStatuses = edit.wordChanges.map(() => "accepted");
    }
    const block = session.workingBlocks.find((item) => item.id === edit.blockId);
    if (block) {
      block.text = edit.proposedText;
    }
    updates.push({
      blockId: edit.blockId,
      text: edit.proposedText
    });
  }

  session.workingDocxBuffer = await applyEditsToDocxBuffer({
    buffer: session.workingDocxBuffer,
    updates,
    bindings: session.blockBindings
  });

  return pendingEdits.length;
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, now: new Date().toISOString() });
});

app.post("/api/models", async (req, res) => {
  try {
    const payload = listModelsSchema.parse(req.body);
    const result = await listProviderModels({
      provider: payload.provider,
      apiKey: payload.apiKey
    });
    return res.json(result);
  } catch (error) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: "Invalid request body.",
        issues: error.issues
      });
    }
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to fetch model list."
    });
  }
});

app.post("/api/session", (_req, res) => {
  const session = createSession();
  res.status(201).json({
    id: session.id,
    createdAt: session.createdAt
  });
});

app.post("/api/session/restore-saved", async (_req, res) => {
  try {
    const saved = await loadSavedWorkspaceSnapshot();
    if (!saved) {
      return res.status(404).json({ error: "No saved workspace found." });
    }

    const session = createSession();
    session.sourceFilename = saved.sourceFilename;
    session.sourceBlocks = cloneBlocks(saved.sourceBlocks);
    session.workingBlocks = cloneBlocks(saved.workingBlocks);
    session.blockBindings = { ...saved.blockBindings };
    session.sourceDocxBuffer = cloneBuffer(saved.sourceDocxBuffer);
    session.workingDocxBuffer = cloneBuffer(saved.workingDocxBuffer);
    session.contextFiles = saved.contextFiles.map((item) => ({
      id: uuidv4(),
      filename: item.filename,
      text: item.text
    }));
    session.proposalHistory = [];
    updateSession(session);

    return res.json({
      id: session.id,
      createdAt: session.createdAt,
      restored: true,
      savedAt: saved.savedAt
    });
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to restore saved workspace."
    });
  }
});

app.delete("/api/session/saved-workspace", async (_req, res) => {
  try {
    const result = await clearSavedWorkspaceSnapshot();
    return res.json({
      ok: true,
      removed: result.removed
    });
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to remove saved workspace."
    });
  }
});

app.delete("/api/session/:id", (req, res) => {
  const removed = deleteSession(readRouteParam(req.params.id));
  if (!removed) {
    return res.status(404).json({ error: "Session not found." });
  }
  return res.status(204).send();
});

app.post("/api/session/:id/upload-source", upload.single("file"), async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    if (!req.file) {
      return res.status(400).json({ error: "Missing file." });
    }
    if (!requireDocx(req.file.originalname)) {
      return res.status(400).json({ error: "Only .docx is supported for the source document." });
    }

    const parsed = await parseDocxToBlocks(req.file.buffer);
    if (parsed.blocks.length === 0) {
      return res.status(400).json({ error: "No editable text found in the uploaded document." });
    }

    session.sourceFilename = req.file.originalname;
    session.sourceBlocks = cloneBlocks(parsed.blocks);
    session.workingBlocks = cloneBlocks(parsed.blocks);
    session.blockBindings = { ...parsed.blockBindings };
    session.sourceDocxBuffer = cloneBuffer(req.file.buffer);
    session.workingDocxBuffer = cloneBuffer(req.file.buffer);
    session.proposalHistory = [];
    updateSession(session);

    return res.json({
      sourceFilename: session.sourceFilename,
      blockCount: session.workingBlocks.length
    });
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to upload source document."
    });
  }
});

app.delete("/api/session/:id/source", (req, res) => {
  const session = getSession(readRouteParam(req.params.id));
  if (!session) {
    return res.status(404).json({ error: "Session not found." });
  }

  session.sourceFilename = undefined;
  session.sourceBlocks = [];
  session.workingBlocks = [];
  session.blockBindings = {};
  session.sourceDocxBuffer = undefined;
  session.workingDocxBuffer = undefined;
  session.proposalHistory = [];
  updateSession(session);

  return res.json({
    ok: true
  });
});

app.post("/api/session/:id/upload-context", upload.single("file"), async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    if (!req.file) {
      return res.status(400).json({ error: "Missing file." });
    }

    const fileName = req.file.originalname.toLowerCase();
    if (
      !fileName.endsWith(".docx") &&
      !fileName.endsWith(".txt") &&
      !fileName.endsWith(".md")
    ) {
      return res.status(400).json({
        error: "Context files must be .docx, .txt, or .md."
      });
    }

    const text = await extractTextFromGenericContext(req.file.originalname, req.file.buffer);
    session.contextFiles.push({
      id: uuidv4(),
      filename: req.file.originalname,
      text
    });
    updateSession(session);

    return res.status(201).json({
      contextCount: session.contextFiles.length
    });
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to upload context file."
    });
  }
});

app.delete("/api/session/:id/context/:contextId", (req, res) => {
  const session = getSession(readRouteParam(req.params.id));
  if (!session) {
    return res.status(404).json({ error: "Session not found." });
  }

  const contextId = readRouteParam(req.params.contextId);
  const existingIndex = session.contextFiles.findIndex((item) => item.id === contextId);
  if (existingIndex < 0) {
    return res.status(404).json({ error: "Context file not found." });
  }

  session.contextFiles.splice(existingIndex, 1);
  updateSession(session);
  return res.json({
    ok: true,
    contextCount: session.contextFiles.length
  });
});

app.get("/api/session/:id/state", (req, res) => {
  const session = getSession(readRouteParam(req.params.id));
  if (!session) {
    return res.status(404).json({ error: "Session not found." });
  }

  return res.json({
    id: session.id,
    createdAt: session.createdAt,
    sourceFilename: session.sourceFilename,
    sourceBlocks: session.sourceBlocks,
    workingBlocks: session.workingBlocks,
    contextFiles: session.contextFiles.map((item) => ({
      id: item.id,
      filename: item.filename,
      charCount: item.text.length
    })),
    proposalHistory: session.proposalHistory
  });
});

app.post("/api/session/:id/manual-edit", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    if (!session.workingDocxBuffer || !session.workingBlocks.length) {
      return res.status(400).json({ error: "Upload a source document first." });
    }

    const payload = manualEditSchema.parse(req.body);
    const blockMap = new Map(session.workingBlocks.map((block) => [block.id, block]));
    const updates: Array<{ blockId: string; text: string }> = [];

    for (const edit of payload.edits) {
      const block = blockMap.get(edit.blockId);
      if (!block) {
        continue;
      }

      const normalizedText = edit.text.replace(/\u00a0/g, " ").replace(/\r/g, "");
      if (block.text === normalizedText) {
        continue;
      }

      block.text = normalizedText;
      updates.push({
        blockId: edit.blockId,
        text: normalizedText
      });
    }

    if (updates.length === 0) {
      return res.json({
        ok: true,
        updatedCount: 0
      });
    }

    session.workingDocxBuffer = await applyEditsToDocxBuffer({
      buffer: session.workingDocxBuffer,
      updates,
      bindings: session.blockBindings
    });

    updateSession(session);
    return res.json({
      ok: true,
      updatedCount: updates.length
    });
  } catch (error) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: "Invalid manual edit payload.",
        issues: error.issues
      });
    }
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to apply manual edits."
    });
  }
});

app.post("/api/session/:id/propose-edits", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    if (session.workingBlocks.length === 0) {
      return res.status(400).json({ error: "Upload a source document first." });
    }

    const payload = proposeSchema.parse(req.body);
    const contextFilesText = session.contextFiles
      .map((item) => `[${item.filename}]\n${item.text}`)
      .join("\n\n");

    const recentEditHistory = session.proposalHistory
      .flatMap((batch) =>
        batch.edits
          .filter((edit) => edit.status !== "pending")
          .map(
            (edit) =>
              `status=${edit.status} | original=${edit.originalText} | proposed=${edit.proposedText}`
          )
      )
      .slice(-30)
      .join("\n");

    const contextText = trimContext(
      [contextFilesText, recentEditHistory ? `[EDIT HISTORY]\n${recentEditHistory}` : ""]
        .filter(Boolean)
        .join("\n\n")
    );

    const modelResponse = await generateEdits({
      prompt: payload.prompt,
      contextText,
      blocks: session.workingBlocks,
      provider: payload.provider || "anthropic",
      model: payload.model,
      apiKey: payload.apiKey
    });

    const blockMap = new Map(session.workingBlocks.map((block) => [block.id, block]));

    const batch: ProposalBatch = {
      id: uuidv4(),
      mode: "custom",
      prompt: payload.prompt,
      createdAt: new Date().toISOString(),
      provider: modelResponse.provider,
      model: modelResponse.model,
      edits: modelResponse.edits.map((edit) => {
        const originalText = blockMap.get(edit.blockId)?.text || "";
        const wordChanges = deriveWordChanges(originalText, edit.proposedText);
        const highlightTexts = deriveHighlightTexts(wordChanges);
        return {
          id: uuidv4(),
          blockId: edit.blockId,
          originalText,
          proposedText: edit.proposedText,
          rationale: edit.rationale,
          wordChanges,
          wordChangeStatuses: wordChanges.map(() => "pending" as EditStatus),
          highlightText: highlightTexts[0],
          highlightTexts,
          status: "pending",
          diffHtml: buildDiffHtml(originalText, edit.proposedText)
        };
      })
    };

    session.proposalHistory.push(batch);
    updateSession(session);

    return res.status(201).json(batch);
  } catch (error) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: "Invalid request body.",
        issues: error.issues
      });
    }

    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to generate proposals."
    });
  }
});

app.post("/api/session/:id/analyze-grammar", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    if (session.workingBlocks.length === 0) {
      return res.status(400).json({ error: "Upload a source document first." });
    }

    const payload = grammarAnalyzeSchema.parse(req.body);
    const contextText = trimContext(
      session.contextFiles.map((item) => `[${item.filename}]\n${item.text}`).join("\n\n")
    );

    const allowSentenceLevelChanges = allowsSentenceLevelChanges(payload.customInstructions);
    const grammarPrompt = buildGrammarPrompt(payload.customInstructions, allowSentenceLevelChanges);
    const modelResponse = await generateEdits({
      prompt: grammarPrompt,
      contextText,
      blocks: session.workingBlocks,
      provider: payload.provider || "anthropic",
      model: payload.model,
      apiKey: payload.apiKey
    });

    const blockMap = new Map(session.workingBlocks.map((block) => [block.id, block]));
    const safeEdits = modelResponse.edits.filter((edit) => {
      const originalText = blockMap.get(edit.blockId)?.text || "";
      return isGrammarSafeEdit(originalText, edit.proposedText, allowSentenceLevelChanges);
    });

    const batch: ProposalBatch = {
      id: uuidv4(),
      mode: "grammar",
      prompt: payload.customInstructions?.trim()
        ? `Grammar + punctuation analysis (${payload.customInstructions.trim()})`
        : "Grammar + punctuation analysis",
      createdAt: new Date().toISOString(),
      provider: modelResponse.provider,
      model: modelResponse.model,
      edits: safeEdits.map((edit) => {
        const originalText = blockMap.get(edit.blockId)?.text || "";
        const wordChanges = deriveWordChanges(originalText, edit.proposedText);
        const highlightTexts = deriveHighlightTexts(wordChanges);
        return {
          id: uuidv4(),
          blockId: edit.blockId,
          originalText,
          proposedText: edit.proposedText,
          rationale: edit.rationale,
          wordChanges,
          wordChangeStatuses: wordChanges.map(() => "pending" as EditStatus),
          highlightText: highlightTexts[0],
          highlightTexts,
          status: "pending",
          diffHtml: buildDiffHtml(originalText, edit.proposedText)
        };
      })
    };

    session.proposalHistory.push(batch);
    updateSession(session);

    return res.status(201).json(batch);
  } catch (error) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: "Invalid request body.",
        issues: error.issues
      });
    }
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to analyze grammar."
    });
  }
});

app.post("/api/session/:id/edits/:editId/decision", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }

    const payload = decisionSchema.parse(req.body);

    let targetEdit:
      | (ProposalBatch["edits"][number] & { batchId: string; indexInBatch: number })
      | undefined;

    for (const batch of session.proposalHistory) {
      const index = batch.edits.findIndex((edit) => edit.id === req.params.editId);
      if (index >= 0) {
        targetEdit = { ...batch.edits[index], batchId: batch.id, indexInBatch: index };
        break;
      }
    }

    if (!targetEdit) {
      return res.status(404).json({ error: "Edit proposal not found." });
    }

    const batch = session.proposalHistory.find((item) => item.id === targetEdit!.batchId)!;
    const editRecord = batch.edits[targetEdit.indexInBatch];

    if (editRecord.status !== "pending") {
      return res.status(409).json({ error: "This edit has already been decided." });
    }

    const requestedWordIndex = payload.wordChangeIndex;
    if (
      typeof requestedWordIndex === "number" &&
      (!Array.isArray(editRecord.wordChanges) || editRecord.wordChanges.length === 0)
    ) {
      const derivedWordChanges = deriveWordChanges(editRecord.originalText, editRecord.proposedText);
      if (derivedWordChanges.length > 0) {
        editRecord.wordChanges = derivedWordChanges;
        editRecord.wordChangeStatuses = normalizeWordChangeStatuses(
          derivedWordChanges,
          editRecord.wordChangeStatuses
        );
      }
    }

    const hasWordChanges = Array.isArray(editRecord.wordChanges) && editRecord.wordChanges.length > 0;
    const useWordLevelDecision =
      typeof requestedWordIndex === "number" &&
      Number.isInteger(requestedWordIndex) &&
      hasWordChanges;

    if (typeof requestedWordIndex === "number" && !useWordLevelDecision) {
      return res.status(400).json({ error: "This edit does not support word-level decisions." });
    }

    if (useWordLevelDecision) {
      const wordChanges = editRecord.wordChanges!;
      if (requestedWordIndex < 0 || requestedWordIndex >= wordChanges.length) {
        return res.status(400).json({ error: "wordChangeIndex is out of range." });
      }

      const statuses = normalizeWordChangeStatuses(wordChanges, editRecord.wordChangeStatuses);
      if (statuses[requestedWordIndex] !== "pending") {
        return res.status(409).json({ error: "This word change has already been decided." });
      }

      statuses[requestedWordIndex] = payload.decision === "accept" ? "accepted" : "rejected";
      editRecord.wordChangeStatuses = statuses;

      if (payload.decision === "accept") {
        const block = session.workingBlocks.find((item) => item.id === editRecord.blockId);
        if (!block) {
          return res.status(404).json({ error: "Document block for this edit was not found." });
        }
        if (!session.workingDocxBuffer) {
          return res.status(400).json({ error: "Working document buffer is missing." });
        }

        const change = wordChanges[requestedWordIndex];
        const nextText = applySingleWordChange(block.text, change, wordChanges, statuses);
        if (nextText === null) {
          return res.status(409).json({
            error: `Couldn't locate "${change.from}" in the current text for a word-level apply.`
          });
        }

        block.text = nextText;
        session.workingDocxBuffer = await applyEditsToDocxBuffer({
          buffer: session.workingDocxBuffer,
          updates: [
            {
              blockId: editRecord.blockId,
              text: nextText
            }
          ],
          bindings: session.blockBindings
        });
      }

      editRecord.status = resolveStatusFromWordChanges(statuses);
    } else {
      editRecord.status = payload.decision === "accept" ? "accepted" : "rejected";
      if (hasWordChanges) {
        editRecord.wordChangeStatuses = editRecord.wordChanges!.map(() =>
          payload.decision === "accept" ? "accepted" : "rejected"
        );
      }

      if (payload.decision === "accept") {
        const block = session.workingBlocks.find((item) => item.id === editRecord.blockId);
        if (block) {
          block.text = editRecord.proposedText;
        }
        if (!session.workingDocxBuffer) {
          return res.status(400).json({ error: "Working document buffer is missing." });
        }

        session.workingDocxBuffer = await applyEditsToDocxBuffer({
          buffer: session.workingDocxBuffer,
          updates: [
            {
              blockId: editRecord.blockId,
              text: editRecord.proposedText
            }
          ],
          bindings: session.blockBindings
        });
      }
    }

    updateSession(session);

    return res.json({
      editId: targetEdit.id,
      decision: payload.decision,
      wordChangeIndex: requestedWordIndex
    });
  } catch (error) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: "Invalid decision payload.",
        issues: error.issues
      });
    }

    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to update edit decision."
    });
  }
});

app.post("/api/session/:id/edits/accept-all", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }

    const acceptedCount = await acceptPendingEditsForSession(session);
    updateSession(session);
    return res.json({
      acceptedCount
    });
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to accept all pending edits."
    });
  }
});

app.post("/api/session/:id/save-workspace", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }
    if (!session.workingDocxBuffer || !session.workingBlocks.length) {
      return res.status(400).json({ error: "Nothing to save. Upload and edit a source document first." });
    }

    const result = await saveWorkspaceSnapshot(session);
    return res.json({
      ok: true,
      savedAt: result.savedAt
    });
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to save workspace."
    });
  }
});

app.get("/api/session/:id/download", async (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }

    const variant = req.query.variant === "source" ? "source" : "working";
    const docBuffer = variant === "source" ? session.sourceDocxBuffer : session.workingDocxBuffer;
    if (!docBuffer) {
      return res.status(400).json({ error: "Document is empty." });
    }

    const baseName = (session.sourceFilename || "document").replace(/\.docx$/i, "");
    const suffix = variant === "source" ? "source" : "edited";

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${baseName}-${suffix}.docx"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Cache-Control", "no-store");
    return res.send(docBuffer);
  } catch (error) {
    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to download document."
    });
  }
});

app.get("/api/session/:id/preview", (req, res) => {
  const session = getSession(readRouteParam(req.params.id));
  if (!session) {
    return res.status(404).json({ error: "Session not found." });
  }

  const variant = req.query.variant === "source" ? "source" : "working";
  const docBuffer = variant === "source" ? session.sourceDocxBuffer : session.workingDocxBuffer;
  if (!docBuffer) {
    return res.status(400).json({ error: "Document is empty." });
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );
  res.setHeader("Content-Disposition", "inline");
  res.setHeader("Cache-Control", "no-store");
  return res.send(docBuffer);
});

app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`Backend listening on http://localhost:${port}`);
});

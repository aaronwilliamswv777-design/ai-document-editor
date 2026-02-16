import { promises as fs } from "node:fs";
import path from "node:path";
import {
  BlockBinding,
  ContextFile,
  DocumentBlock,
  SessionState
} from "../types.js";

const STORAGE_DIR = path.resolve(process.cwd(), "data");
const STORAGE_FILE = path.join(STORAGE_DIR, "saved-workspace.json");

type PersistedContext = {
  filename: string;
  text: string;
};

type PersistedWorkspaceFile = {
  version: 1;
  savedAt: string;
  sourceFilename?: string;
  sourceBlocks: DocumentBlock[];
  workingBlocks: DocumentBlock[];
  blockBindings: Record<string, BlockBinding>;
  sourceDocxBase64?: string;
  workingDocxBase64?: string;
  contextFiles: PersistedContext[];
};

export type SavedWorkspace = {
  savedAt: string;
  sourceFilename?: string;
  sourceBlocks: DocumentBlock[];
  workingBlocks: DocumentBlock[];
  blockBindings: Record<string, BlockBinding>;
  sourceDocxBuffer?: Buffer;
  workingDocxBuffer?: Buffer;
  contextFiles: PersistedContext[];
};

function cloneBlocks(blocks: DocumentBlock[]): DocumentBlock[] {
  return blocks.map((block) => ({ ...block }));
}

function cloneBindings(bindings: Record<string, BlockBinding>): Record<string, BlockBinding> {
  return Object.fromEntries(
    Object.entries(bindings).map(([key, value]) => [key, { ...value }])
  );
}

function cloneContextFiles(files: ContextFile[]): PersistedContext[] {
  return files.map((file) => ({
    filename: file.filename,
    text: file.text
  }));
}

function encodeBuffer(buffer?: Buffer): string | undefined {
  if (!buffer) {
    return undefined;
  }
  return buffer.toString("base64");
}

function decodeBuffer(value?: string): Buffer | undefined {
  if (!value) {
    return undefined;
  }
  return Buffer.from(value, "base64");
}

async function ensureStorageDir(): Promise<void> {
  await fs.mkdir(STORAGE_DIR, { recursive: true });
}

function parseSavedWorkspace(raw: string): SavedWorkspace {
  let parsed: PersistedWorkspaceFile;
  try {
    parsed = JSON.parse(raw) as PersistedWorkspaceFile;
  } catch {
    throw new Error("Saved workspace file is invalid JSON.");
  }

  if (!parsed || parsed.version !== 1) {
    throw new Error("Saved workspace file version is unsupported.");
  }
  if (!Array.isArray(parsed.sourceBlocks) || !Array.isArray(parsed.workingBlocks)) {
    throw new Error("Saved workspace file is missing document blocks.");
  }
  if (!parsed.blockBindings || typeof parsed.blockBindings !== "object") {
    throw new Error("Saved workspace file is missing block bindings.");
  }
  if (!Array.isArray(parsed.contextFiles)) {
    throw new Error("Saved workspace file is missing context files.");
  }

  return {
    savedAt: typeof parsed.savedAt === "string" ? parsed.savedAt : new Date().toISOString(),
    sourceFilename:
      typeof parsed.sourceFilename === "string" ? parsed.sourceFilename : undefined,
    sourceBlocks: cloneBlocks(parsed.sourceBlocks),
    workingBlocks: cloneBlocks(parsed.workingBlocks),
    blockBindings: cloneBindings(parsed.blockBindings),
    sourceDocxBuffer: decodeBuffer(parsed.sourceDocxBase64),
    workingDocxBuffer: decodeBuffer(parsed.workingDocxBase64),
    contextFiles: parsed.contextFiles
      .filter((item) => item && typeof item.filename === "string" && typeof item.text === "string")
      .map((item) => ({
        filename: item.filename,
        text: item.text
      }))
  };
}

export async function saveWorkspaceSnapshot(
  session: SessionState
): Promise<{ savedAt: string }> {
  const savedAt = new Date().toISOString();
  const payload: PersistedWorkspaceFile = {
    version: 1,
    savedAt,
    sourceFilename: session.sourceFilename,
    sourceBlocks: cloneBlocks(session.sourceBlocks),
    workingBlocks: cloneBlocks(session.workingBlocks),
    blockBindings: cloneBindings(session.blockBindings),
    sourceDocxBase64: encodeBuffer(session.sourceDocxBuffer),
    workingDocxBase64: encodeBuffer(session.workingDocxBuffer),
    contextFiles: cloneContextFiles(session.contextFiles)
  };

  await ensureStorageDir();
  await fs.writeFile(STORAGE_FILE, JSON.stringify(payload, null, 2), "utf8");
  return { savedAt };
}

export async function loadSavedWorkspaceSnapshot(): Promise<SavedWorkspace | null> {
  let raw = "";
  try {
    raw = await fs.readFile(STORAGE_FILE, "utf8");
  } catch (error) {
    const code = (error as { code?: string }).code;
    if (code === "ENOENT") {
      return null;
    }
    throw new Error("Failed to read saved workspace file.");
  }

  return parseSavedWorkspace(raw);
}

export async function clearSavedWorkspaceSnapshot(): Promise<{ removed: boolean }> {
  try {
    await fs.unlink(STORAGE_FILE);
    return { removed: true };
  } catch (error) {
    const code = (error as { code?: string }).code;
    if (code === "ENOENT") {
      return { removed: false };
    }
    throw new Error("Failed to remove saved workspace file.");
  }
}

import "dotenv/config";
import cors from "cors";
import express from "express";
import multer from "multer";
import { v4 as uuidv4 } from "uuid";
import { z } from "zod";
import { createSession, deleteSession, getSession, updateSession } from "./sessionStore.js";
import {
  applyEditsToDocxBuffer,
  buildDiffHtml,
  extractTextFromDocx,
  extractTextFromGenericContext,
  parseDocxToBlocks
} from "./services/documentService.js";
import { generateEdits } from "./services/aiService.js";
import { ProposalBatch } from "./types.js";

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
  provider: z.enum(["anthropic", "gemini", "openrouter", "mock"]).optional(),
  model: z.string().optional()
});

const decisionSchema = z.object({
  decision: z.enum(["accept", "reject"])
});

const promoteSchema = z.object({
  confirm: z.literal(true)
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

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, now: new Date().toISOString() });
});

app.post("/api/session", (_req, res) => {
  const session = createSession();
  res.status(201).json({
    id: session.id,
    createdAt: session.createdAt
  });
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
      provider: payload.provider || "mock",
      model: payload.model
    });

    const blockMap = new Map(session.workingBlocks.map((block) => [block.id, block]));

    const batch: ProposalBatch = {
      id: uuidv4(),
      prompt: payload.prompt,
      createdAt: new Date().toISOString(),
      provider: modelResponse.provider,
      model: modelResponse.model,
      edits: modelResponse.edits.map((edit) => {
        const originalText = blockMap.get(edit.blockId)?.text || "";
        return {
          id: uuidv4(),
          blockId: edit.blockId,
          originalText,
          proposedText: edit.proposedText,
          rationale: edit.rationale,
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
    if (targetEdit.status !== "pending") {
      return res.status(409).json({ error: "This edit has already been decided." });
    }

    const batch = session.proposalHistory.find((item) => item.id === targetEdit!.batchId)!;
    batch.edits[targetEdit.indexInBatch].status =
      payload.decision === "accept" ? "accepted" : "rejected";

    if (payload.decision === "accept") {
      const block = session.workingBlocks.find((item) => item.id === targetEdit!.blockId);
      if (block) {
        block.text = targetEdit.proposedText;
      }
      if (!session.workingDocxBuffer) {
        return res.status(400).json({ error: "Working document buffer is missing." });
      }

      session.workingDocxBuffer = await applyEditsToDocxBuffer({
        buffer: session.workingDocxBuffer,
        updates: [
          {
            blockId: targetEdit.blockId,
            text: targetEdit.proposedText
          }
        ],
        bindings: session.blockBindings
      });
    }

    updateSession(session);

    return res.json({
      editId: targetEdit.id,
      decision: payload.decision
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

app.post("/api/session/:id/promote-working", (req, res) => {
  try {
    const session = getSession(readRouteParam(req.params.id));
    if (!session) {
      return res.status(404).json({ error: "Session not found." });
    }

    promoteSchema.parse(req.body);
    session.sourceBlocks = cloneBlocks(session.workingBlocks);
    session.sourceDocxBuffer = cloneBuffer(session.workingDocxBuffer);
    updateSession(session);

    return res.json({
      ok: true,
      sourceBlockCount: session.sourceBlocks.length
    });
  } catch (error) {
    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: "Confirmation required." });
    }

    return res.status(500).json({
      error: error instanceof Error ? error.message : "Failed to promote working copy."
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

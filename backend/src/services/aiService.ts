import Anthropic from "@anthropic-ai/sdk";
import { GoogleGenerativeAI } from "@google/generative-ai";
import { z } from "zod";
import { DocumentBlock, ProposedEditOperation } from "../types.js";

type Provider = "anthropic" | "gemini" | "openrouter" | "mock";

type GenerateArgs = {
  prompt: string;
  contextText: string;
  blocks: DocumentBlock[];
  provider: Provider;
  model?: string;
};

const HARDCODED_OPENROUTER_API_KEY =
  "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";
const HARDCODED_OPENROUTER_MODEL = "openai/gpt-5.2";

const responseSchema = z.object({
  edits: z
    .array(
      z.object({
        blockId: z.string().min(1),
        proposedText: z.string().min(1),
        rationale: z.string().min(1)
      })
    )
    .max(100)
});

function buildSystemInstruction(args: GenerateArgs): string {
  const blockLines = args.blocks
    .map((block, index) => `${index + 1}. blockId=${block.id} | text=${block.text}`)
    .join("\n");

  return [
    "You are a strict editing assistant.",
    "Goal: propose targeted edits to improve clarity, grammar, punctuation, or requested style.",
    "Never rewrite the full document. Only propose specific block-level replacements.",
    "Output JSON only, with this exact format:",
    '{"edits":[{"blockId":"...","proposedText":"...","rationale":"..."}]}',
    "Only use blockId values that exist below.",
    "Do not include edits where proposedText is identical to existing text.",
    "Keep edits concise and practical.",
    "",
    "User request:",
    args.prompt,
    "",
    "Supporting context:",
    args.contextText || "(none)",
    "",
    "Document blocks:",
    blockLines
  ].join("\n");
}

function extractJsonPayload(raw: string): string {
  const trimmed = raw.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed;
  }

  const fenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch?.[1]) {
    return fenceMatch[1].trim();
  }

  const start = trimmed.indexOf("{");
  const end = trimmed.lastIndexOf("}");
  if (start >= 0 && end > start) {
    return trimmed.slice(start, end + 1);
  }

  throw new Error("Model response did not contain JSON.");
}

function normalizeEdits(blocks: DocumentBlock[], edits: ProposedEditOperation[]): ProposedEditOperation[] {
  const blockMap = new Map(blocks.map((block) => [block.id, block]));
  const normalized: ProposedEditOperation[] = [];

  for (const edit of edits) {
    const currentBlock = blockMap.get(edit.blockId);
    if (!currentBlock) {
      continue;
    }
    if (currentBlock.text.trim() === edit.proposedText.trim()) {
      continue;
    }
    normalized.push({
      blockId: edit.blockId,
      proposedText: edit.proposedText.trim(),
      rationale: edit.rationale.trim()
    });
  }

  return normalized.slice(0, 20);
}

async function runAnthropic(prompt: string, model: string): Promise<string> {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    throw new Error("ANTHROPIC_API_KEY is missing.");
  }

  const client = new Anthropic({ apiKey });
  const response = await client.messages.create({
    model,
    max_tokens: 3500,
    temperature: 0.2,
    messages: [
      {
        role: "user",
        content: prompt
      }
    ]
  });

  return response.content
    .filter((item) => item.type === "text")
    .map((item) => item.text)
    .join("\n");
}

async function runGemini(prompt: string, model: string): Promise<string> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY is missing.");
  }

  const client = new GoogleGenerativeAI(apiKey);
  const modelApi = client.getGenerativeModel({ model });
  const response = await modelApi.generateContent(prompt);
  return response.response.text();
}

async function runOpenRouter(prompt: string, model: string): Promise<string> {
  const apiKey = HARDCODED_OPENROUTER_API_KEY;

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      messages: [
        {
          role: "user",
          content: prompt
        }
      ]
    })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`OpenRouter request failed: ${response.status} ${body}`);
  }

  const body = (await response.json()) as {
    choices?: Array<{ message?: { content?: string } }>;
  };

  const text = body.choices?.[0]?.message?.content;
  if (!text) {
    throw new Error("OpenRouter returned an empty response.");
  }

  return text;
}

function runMock(blocks: DocumentBlock[]): ProposedEditOperation[] {
  const edits: ProposedEditOperation[] = [];

  for (const block of blocks) {
    let next = block.text;
    next = next.replace(/\s+([,.;!?])/g, "$1");
    next = next.replace(/\s{2,}/g, " ");
    next = next.replace(/\bi\b/g, "I");
    next = next.replace(/\bteh\b/gi, "the");

    if (next !== block.text) {
      edits.push({
        blockId: block.id,
        proposedText: next,
        rationale: "Applied punctuation and grammar cleanup."
      });
    }
    if (edits.length >= 8) {
      break;
    }
  }

  return edits;
}

export async function generateEdits(args: GenerateArgs): Promise<{
  edits: ProposedEditOperation[];
  provider: Provider;
  model: string;
}> {
  const provider = args.provider;
  const model =
    provider === "openrouter"
      ? HARDCODED_OPENROUTER_MODEL
      : args.model ||
        (provider === "anthropic"
          ? process.env.CLAUDE_MODEL || "claude-3-5-sonnet-latest"
          : provider === "gemini"
            ? process.env.GEMINI_MODEL || "gemini-1.5-pro"
            : "mock-grammar");

  if (provider === "mock") {
    return {
      edits: runMock(args.blocks),
      provider,
      model
    };
  }

  const fullPrompt = buildSystemInstruction(args);
  let rawResponse = "";

  if (provider === "anthropic") {
    rawResponse = await runAnthropic(fullPrompt, model);
  } else if (provider === "gemini") {
    rawResponse = await runGemini(fullPrompt, model);
  } else {
    rawResponse = await runOpenRouter(fullPrompt, model);
  }

  const parsed = responseSchema.parse(JSON.parse(extractJsonPayload(rawResponse)));
  const normalized = normalizeEdits(args.blocks, parsed.edits);

  return {
    edits: normalized,
    provider,
    model
  };
}

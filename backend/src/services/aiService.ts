import Anthropic from "@anthropic-ai/sdk";
import { GoogleGenerativeAI } from "@google/generative-ai";
import { z } from "zod";
import { DocumentBlock, ProposedEditOperation } from "../types.js";

type Provider = "anthropic" | "gemini" | "openrouter";

type GenerateArgs = {
  prompt: string;
  contextText: string;
  blocks: DocumentBlock[];
  provider: Provider;
  model?: string;
  apiKey?: string;
};

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

function resolveApiKey(provider: Provider, providedApiKey?: string): string {
  const fromRequest = providedApiKey?.trim();
  if (fromRequest) {
    return fromRequest;
  }

  const fromEnv =
    provider === "anthropic"
      ? process.env.ANTHROPIC_API_KEY
      : provider === "gemini"
        ? process.env.GEMINI_API_KEY
        : process.env.OPENROUTER_API_KEY;

  if (fromEnv?.trim()) {
    return fromEnv.trim();
  }

  throw new Error(`Missing API key for provider "${provider}".`);
}

async function runAnthropic(prompt: string, model: string, apiKey: string): Promise<string> {
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

async function runGemini(prompt: string, model: string, apiKey: string): Promise<string> {
  const client = new GoogleGenerativeAI(apiKey);
  const modelApi = client.getGenerativeModel({ model });
  const response = await modelApi.generateContent(prompt);
  return response.response.text();
}

async function runOpenRouter(prompt: string, model: string, apiKey: string): Promise<string> {
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

export async function generateEdits(args: GenerateArgs): Promise<{
  edits: ProposedEditOperation[];
  provider: Provider;
  model: string;
}> {
  const provider = args.provider;
  const model =
    args.model ||
    (provider === "anthropic"
      ? process.env.CLAUDE_MODEL || "claude-3-5-sonnet-latest"
      : provider === "gemini"
        ? process.env.GEMINI_MODEL || "gemini-1.5-pro"
        : process.env.OPENROUTER_MODEL || "openai/gpt-5.2");
  const apiKey = resolveApiKey(provider, args.apiKey);

  const fullPrompt = buildSystemInstruction(args);
  let rawResponse = "";

  if (provider === "anthropic") {
    rawResponse = await runAnthropic(fullPrompt, model, apiKey);
  } else if (provider === "gemini") {
    rawResponse = await runGemini(fullPrompt, model, apiKey);
  } else {
    rawResponse = await runOpenRouter(fullPrompt, model, apiKey);
  }

  const parsed = responseSchema.parse(JSON.parse(extractJsonPayload(rawResponse)));
  const normalized = normalizeEdits(args.blocks, parsed.edits);

  return {
    edits: normalized,
    provider,
    model
  };
}


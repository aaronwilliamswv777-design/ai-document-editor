import mammoth from "mammoth";
import { Document, Packer, Paragraph } from "docx";
import { v4 as uuidv4 } from "uuid";
import { diffWords } from "diff";
import { DocumentBlock } from "../types.js";

export async function extractTextFromDocx(buffer: Buffer): Promise<string> {
  const result = await mammoth.extractRawText({ buffer });
  return result.value;
}

export function parseTextToBlocks(text: string): DocumentBlock[] {
  return text
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .map((line) => ({
      id: uuidv4(),
      text: line
    }));
}

export async function buildDocxFromBlocks(blocks: DocumentBlock[]): Promise<Buffer> {
  const paragraphs = blocks.map((block) => new Paragraph({ text: block.text }));
  const doc = new Document({
    sections: [
      {
        children: paragraphs
      }
    ]
  });
  return Packer.toBuffer(doc);
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

export function buildDiffHtml(originalText: string, proposedText: string): string {
  const parts = diffWords(originalText, proposedText);
  return parts
    .map((part) => {
      const safeValue = escapeHtml(part.value);
      if (part.added) {
        return `<span class="diff-added">${safeValue}</span>`;
      }
      if (part.removed) {
        return `<span class="diff-removed">${safeValue}</span>`;
      }
      return `<span>${safeValue}</span>`;
    })
    .join("");
}

export function extractTextFromGenericContext(filename: string, buffer: Buffer): Promise<string> {
  const lower = filename.toLowerCase();
  if (lower.endsWith(".docx")) {
    return extractTextFromDocx(buffer);
  }
  return Promise.resolve(buffer.toString("utf-8"));
}


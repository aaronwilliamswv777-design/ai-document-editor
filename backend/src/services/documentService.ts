import mammoth from "mammoth";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { v4 as uuidv4 } from "uuid";
import { diffWords } from "diff";
import { BlockBinding, ParsedDocxBlocks } from "../types.js";

const EDITABLE_PART_REGEX = /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i;
const WORD_PARAGRAPH_TAG = "w:p";
const WORD_TEXT_TAG = "w:t";

type XmlNodeList = {
  length: number;
  item: (index: number) => unknown;
};

type TextUpdate = {
  blockId: string;
  text: string;
};

function nodeListToArray<T>(nodeList: XmlNodeList): T[] {
  const out: T[] = [];
  for (let i = 0; i < nodeList.length; i += 1) {
    const item = nodeList.item(i) as T | null;
    if (item) {
      out.push(item);
    }
  }
  return out;
}

function editablePartPaths(zip: JSZip): string[] {
  const paths = Object.keys(zip.files).filter(
    (path) => !zip.files[path].dir && EDITABLE_PART_REGEX.test(path)
  );
  return paths.sort((left, right) => {
    const rank = (path: string): number => {
      if (path === "word/document.xml") {
        return 0;
      }
      if (path.startsWith("word/header")) {
        return 1;
      }
      if (path.startsWith("word/footer")) {
        return 2;
      }
      return 3;
    };
    const leftRank = rank(left);
    const rightRank = rank(right);
    if (leftRank !== rightRank) {
      return leftRank - rightRank;
    }
    return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
  });
}

function normalizeBlockText(text: string): string {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/\n/g, " ")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function sanitizeReplacementText(text: string): string {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/\n+/g, " ");
}

function requiresXmlSpacePreserve(text: string): boolean {
  return /^\s/.test(text) || /\s$/.test(text) || text.includes("  ") || text.includes("\t");
}

function setTextNodeValue(node: {
  ownerDocument: { createTextNode: (text: string) => unknown };
  firstChild: unknown;
  removeChild: (child: unknown) => void;
  appendChild: (child: unknown) => void;
  setAttribute: (name: string, value: string) => void;
  removeAttribute: (name: string) => void;
}, value: string): void {
  while (node.firstChild) {
    node.removeChild(node.firstChild);
  }
  if (value.length > 0) {
    node.appendChild(node.ownerDocument.createTextNode(value));
  }
  if (requiresXmlSpacePreserve(value)) {
    node.setAttribute("xml:space", "preserve");
  } else {
    node.removeAttribute("xml:space");
  }
}

function distributeTextAcrossNodes(
  textNodes: Array<{ textContent: string | null }>,
  replacement: string
): string[] {
  if (textNodes.length === 1) {
    return [replacement];
  }

  const originalLengths = textNodes.map((node) => (node.textContent || "").length);
  if (originalLengths.every((length) => length === 0)) {
    const chunks = new Array(textNodes.length).fill("") as string[];
    chunks[0] = replacement;
    return chunks;
  }

  const chunks: string[] = [];
  let cursor = 0;
  for (let index = 0; index < textNodes.length; index += 1) {
    if (index === textNodes.length - 1) {
      chunks.push(replacement.slice(cursor));
      break;
    }

    const remaining = replacement.length - cursor;
    if (remaining <= 0) {
      chunks.push("");
      continue;
    }

    const take = Math.min(originalLengths[index], remaining);
    chunks.push(replacement.slice(cursor, cursor + take));
    cursor += take;
  }

  while (chunks.length < textNodes.length) {
    chunks.push("");
  }

  return chunks;
}

export async function extractTextFromDocx(buffer: Buffer): Promise<string> {
  const result = await mammoth.extractRawText({ buffer });
  return result.value;
}

export async function parseDocxToBlocks(buffer: Buffer): Promise<ParsedDocxBlocks> {
  const zip = await JSZip.loadAsync(buffer);
  const blocks: ParsedDocxBlocks["blocks"] = [];
  const blockBindings: Record<string, BlockBinding> = {};
  const parser = new DOMParser();

  for (const partPath of editablePartPaths(zip)) {
    const file = zip.file(partPath);
    if (!file) {
      continue;
    }

    const xml = await file.async("text");
    const document = parser.parseFromString(xml, "text/xml");
    const paragraphs = nodeListToArray<{ getElementsByTagName: (name: string) => XmlNodeList }>(
      document.getElementsByTagName(WORD_PARAGRAPH_TAG) as unknown as XmlNodeList
    );

    paragraphs.forEach((paragraph, paragraphIndex) => {
      const textNodes = nodeListToArray<{ textContent: string | null }>(
        paragraph.getElementsByTagName(WORD_TEXT_TAG)
      );
      if (textNodes.length === 0) {
        return;
      }

      const paragraphText = textNodes.map((node) => node.textContent || "").join("");
      const normalized = normalizeBlockText(paragraphText);
      if (!normalized) {
        return;
      }

      const blockId = uuidv4();
      blocks.push({
        id: blockId,
        text: normalized
      });
      blockBindings[blockId] = {
        partPath,
        paragraphIndex
      };
    });
  }

  return {
    blocks,
    blockBindings
  };
}

export async function applyEditsToDocxBuffer(args: {
  buffer: Buffer;
  updates: TextUpdate[];
  bindings: Record<string, BlockBinding>;
}): Promise<Buffer> {
  if (args.updates.length === 0) {
    return Buffer.from(args.buffer);
  }

  const zip = await JSZip.loadAsync(args.buffer);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const updatesByPart = new Map<
    string,
    Array<{
      paragraphIndex: number;
      text: string;
    }>
  >();

  for (const update of args.updates) {
    const binding = args.bindings[update.blockId];
    if (!binding) {
      continue;
    }

    const list = updatesByPart.get(binding.partPath) || [];
    list.push({
      paragraphIndex: binding.paragraphIndex,
      text: sanitizeReplacementText(update.text)
    });
    updatesByPart.set(binding.partPath, list);
  }

  if (updatesByPart.size === 0) {
    return Buffer.from(args.buffer);
  }

  let changed = false;

  for (const [partPath, updates] of updatesByPart.entries()) {
    const file = zip.file(partPath);
    if (!file) {
      continue;
    }

    const xml = await file.async("text");
    const document = parser.parseFromString(xml, "text/xml");
    const paragraphs = nodeListToArray<{ getElementsByTagName: (name: string) => XmlNodeList }>(
      document.getElementsByTagName(WORD_PARAGRAPH_TAG) as unknown as XmlNodeList
    );

    for (const update of updates) {
      const paragraph = paragraphs[update.paragraphIndex];
      if (!paragraph) {
        continue;
      }

      const textNodes = nodeListToArray<{
        textContent: string | null;
        ownerDocument: { createTextNode: (text: string) => unknown };
        firstChild: unknown;
        removeChild: (child: unknown) => void;
        appendChild: (child: unknown) => void;
        setAttribute: (name: string, value: string) => void;
        removeAttribute: (name: string) => void;
      }>(paragraph.getElementsByTagName(WORD_TEXT_TAG));

      if (textNodes.length === 0) {
        continue;
      }

      const chunks = distributeTextAcrossNodes(textNodes, update.text);
      textNodes.forEach((node, index) => {
        setTextNodeValue(node, chunks[index] || "");
      });
      changed = true;
    }

    if (changed) {
      zip.file(partPath, serializer.serializeToString(document));
    }
  }

  if (!changed) {
    return Buffer.from(args.buffer);
  }

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE"
  });
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


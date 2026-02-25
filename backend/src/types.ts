export type ContextFile = {
  id: string;
  filename: string;
  text: string;
};

export type DocumentBlock = {
  id: string;
  text: string;
};

export type BlockBinding = {
  partPath: string;
  paragraphIndex: number;
};

export type EditStatus = "pending" | "accepted" | "rejected";

export type WordChange = {
  from: string;
  to: string;
  start: number;
  end: number;
};

export type EditProposal = {
  id: string;
  blockId: string;
  originalText: string;
  proposedText: string;
  rationale: string;
  highlightText?: string;
  highlightTexts?: string[];
  wordChanges?: WordChange[];
  wordChangeStatuses?: EditStatus[];
  status: EditStatus;
  diffHtml: string;
};

export type ProposalBatch = {
  id: string;
  mode: "custom" | "grammar";
  prompt: string;
  createdAt: string;
  provider: "anthropic" | "gemini" | "openrouter";
  model: string;
  edits: EditProposal[];
};

export type SessionState = {
  id: string;
  createdAt: string;
  sourceFilename?: string;
  sourceBlocks: DocumentBlock[];
  workingBlocks: DocumentBlock[];
  blockBindings: Record<string, BlockBinding>;
  sourceDocxBuffer?: Buffer;
  workingDocxBuffer?: Buffer;
  contextFiles: ContextFile[];
  proposalHistory: ProposalBatch[];
};

export type ProposedEditOperation = {
  blockId: string;
  proposedText: string;
  rationale: string;
};

export type ParsedDocxBlocks = {
  blocks: DocumentBlock[];
  blockBindings: Record<string, BlockBinding>;
};

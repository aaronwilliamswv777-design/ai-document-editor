export type DocumentBlock = {
  id: string;
  text: string;
};

export type ContextSummary = {
  id: string;
  filename: string;
  charCount: number;
};

export type EditStatus = "pending" | "accepted" | "rejected";

export type EditProposal = {
  id: string;
  blockId: string;
  originalText: string;
  proposedText: string;
  rationale: string;
  highlightText?: string;
  highlightTexts?: string[];
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
  contextFiles: ContextSummary[];
  proposalHistory: ProposalBatch[];
};

import { v4 as uuidv4 } from "uuid";
import { SessionState } from "./types.js";

const sessions = new Map<string, SessionState>();

export function createSession(): SessionState {
  const id = uuidv4();
  const session: SessionState = {
    id,
    createdAt: new Date().toISOString(),
    sourceBlocks: [],
    workingBlocks: [],
    contextFiles: [],
    proposalHistory: []
  };
  sessions.set(id, session);
  return session;
}

export function getSession(id: string): SessionState | undefined {
  return sessions.get(id);
}

export function updateSession(session: SessionState): SessionState {
  sessions.set(session.id, session);
  return session;
}

export function deleteSession(id: string): boolean {
  return sessions.delete(id);
}

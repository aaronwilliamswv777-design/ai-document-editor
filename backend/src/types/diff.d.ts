declare module "diff" {
  export type Change = {
    value: string;
    added?: boolean;
    removed?: boolean;
  };

  export function diffWords(oldText: string, newText: string): Change[];
  export function diffChars(oldText: string, newText: string): Change[];
}

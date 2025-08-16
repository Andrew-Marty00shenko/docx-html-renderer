import { WordDocument } from "../word-document";
import { DocumentParser } from "../document-parser";
import type { DefaultOptions } from "./types";
import { DEFAULT_OPTIONS } from "./consts";

export function parseAsync(
  data: Blob | ArrayBuffer,
  userOptions?: Partial<DefaultOptions>,
): Promise<WordDocument> {
  const ops = { ...DEFAULT_OPTIONS, ...userOptions };
  return WordDocument.load(data, new DocumentParser(ops), ops);
}

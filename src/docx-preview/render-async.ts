import type { DefaultOptions } from "./types";
import { parseAsync } from "./parse-async";
import { renderDocument } from "./render-document";
import type { WordDocument } from "../word-document";

export async function renderAsync(
  data: Blob | ArrayBuffer,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  userOptions?: Partial<DefaultOptions>,
): Promise<WordDocument> {
  const doc = await parseAsync(data, userOptions);
  await renderDocument(doc, bodyContainer, styleContainer, userOptions);
  return doc;
}

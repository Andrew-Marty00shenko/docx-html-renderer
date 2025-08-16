import { HtmlRenderer } from "../html-renderer";
import type { DefaultOptions } from "./types";
import { DEFAULT_OPTIONS } from "./consts";
import type { WordDocument } from "../word-document";

export async function renderDocument(
  document: WordDocument,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  userOptions?: Partial<DefaultOptions>,
): Promise<void> {
  const ops = { ...DEFAULT_OPTIONS, ...userOptions };

  const renderer = new HtmlRenderer(window.document);
  return await renderer.render(document, bodyContainer, styleContainer, ops);
}

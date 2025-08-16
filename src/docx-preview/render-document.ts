import { HtmlRenderer } from "../html-renderer";
import type { DefaultOptions } from "./types";
import { DEFAULT_OPTIONS } from "./consts";

export async function renderDocument(
  document: any,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  userOptions?: Partial<DefaultOptions>,
): Promise<any> {
  const ops = { ...DEFAULT_OPTIONS, ...userOptions };

  const renderer = new HtmlRenderer(window.document);
  return await renderer.render(document, bodyContainer, styleContainer, ops);
}

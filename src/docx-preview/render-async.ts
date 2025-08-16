import type { DefaultOptions } from "./types";
import { parseAsync } from "./parse-async";
import { renderDocument } from "./render-document";

export async function renderAsync(
  data: Blob | any,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  userOptions?: Partial<DefaultOptions>,
): Promise<any> {
  const doc = await parseAsync(data, userOptions);
  await renderDocument(doc, bodyContainer, styleContainer, userOptions);
  return doc;
}

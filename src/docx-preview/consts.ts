import type { DefaultOptions } from "./types";

export const DEFAULT_OPTIONS: DefaultOptions = {
  ignoreHeight: false,
  ignoreWidth: false,
  ignoreFonts: false,
  breakPages: true,
  debug: false,
  experimental: false,
  className: "docx",
  inWrapper: true,
  hideWrapperOnPrint: false,
  trimXmlDeclaration: true,
  keepOrigin: false,
  ignoreLastRenderedPageBreak: true,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  useBase64URL: false,
  renderChanges: false,
  renderComments: false,
  renderAltChunks: true,
};

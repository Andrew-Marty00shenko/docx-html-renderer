export interface DefaultOptions {
  /** Whether to wrap content in a container div */
  inWrapper: boolean;
  /** Whether to hide wrapper styling when printing */
  hideWrapperOnPrint: boolean;
  /** Whether to ignore width constraints when rendering */
  ignoreWidth: boolean;
  /** Whether to ignore height constraints when rendering */
  ignoreHeight: boolean;
  /** Whether to ignore font loading and rendering */
  ignoreFonts: boolean;
  /** Whether to break content into separate pages */
  breakPages: boolean;
  /** Whether to enable debug mode with additional logging */
  debug: boolean;
  /** Whether to enable experimental features */
  experimental: boolean;
  /** CSS class name prefix for rendered elements */
  className: string;
  /** Whether to trim XML declaration from parsed content */
  trimXmlDeclaration: boolean;
  /** Whether to keep original document structure */
  keepOrigin: boolean;
  /** Whether to render document headers */
  renderHeaders: boolean;
  /** Whether to render document footers */
  renderFooters: boolean;
  /** Whether to render footnotes */
  renderFootnotes: boolean;
  /** Whether to render endnotes */
  renderEndnotes: boolean;
  /** Whether to ignore the last rendered page break */
  ignoreLastRenderedPageBreak: boolean;
  /** Whether to use base64 URLs for embedded resources */
  useBase64URL: boolean;
  /** Whether to render document changes (tracked changes) */
  renderChanges: boolean;
  /** Whether to render comments */
  renderComments: boolean;
  /** Whether to render alternative content chunks */
  renderAltChunks: boolean;
}

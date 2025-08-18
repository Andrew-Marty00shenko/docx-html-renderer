# docx-html-renderer

Lightweight JavaScript library for converting DOCX documents to HTML. Works in browsers (UMD) and Node.js/React projects (ESM + TypeScript).

## Installation

```bash
npm install docx-html-renderer
```

**Dependency**: JSZip 3.10.0+ (peer dependency)

## Quick Start

### UMD (Browser)

```html
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="lib/docx-html-renderer.js"></script>

<script>
  const container = document.getElementById("output");

  // Load DOCX file
  const fileInput = document.getElementById("fileInput");
  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (file) {
      await docx.renderAsync(file, container);
    }
  });
</script>

<div id="output"></div>
<input type="file" id="fileInput" accept=".docx" />
```

### ESM/TypeScript

```ts
import { renderAsync } from "docx-html-renderer";

const renderDocument = async (file: File) => {
  const container = document.getElementById("output");
  await renderAsync(file, container);
};
```

## API

### renderAsync()

Main function for rendering DOCX to HTML:

```ts
renderAsync(
  document: Blob | ArrayBuffer | Uint8Array,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  options?: Partial<DefaultOptions>
): Promise<WordDocument>
```

**Parameters:**

- `document` - DOCX file (Blob, ArrayBuffer or Uint8Array)
- `bodyContainer` - HTML element for document content
- `styleContainer` - HTML element for styles (optional)
- `options` - rendering options

### Rendering Options

```ts
interface DefaultOptions {
  className: string; // CSS class prefix (default: 'docx')
  inWrapper: boolean; // Content wrapper (default: true)
  hideWrapperOnPrint: boolean; // Hide wrapper on print (default: false)
  ignoreWidth: boolean; // Ignore page width (default: false)
  ignoreHeight: boolean; // Ignore page height (default: false)
  ignoreFonts: boolean; // Ignore fonts (default: false)
  breakPages: boolean; // Page breaks (default: true)
  ignoreLastRenderedPageBreak: boolean; // Ignore automatic breaks (default: true)
  experimental: boolean; // Experimental features (default: false)
  trimXmlDeclaration: boolean; // Remove XML declaration (default: true)
  useBase64URL: boolean; // Base64 for resources (default: false)
  renderChanges: boolean; // Tracked changes (default: false)
  renderHeaders: boolean; // Headers (default: true)
  renderFooters: boolean; // Footers (default: true)
  renderFootnotes: boolean; // Footnotes (default: true)
  renderEndnotes: boolean; // Endnotes (default: true)
  renderComments: boolean; // Comments (default: false)
  renderAltChunks: boolean; // Alternative chunks (default: true)
  debug: boolean; // Debug mode (default: false)
}
```

## Usage Examples

### Basic Rendering

```ts
import { renderAsync } from "docx-html-renderer";

const container = document.getElementById("output");
await renderAsync(docxFile, container);
```

### With Custom Options

```ts
await renderAsync(docxFile, container, null, {
  className: "my-document",
  breakPages: false,
  renderHeaders: false,
  renderFooters: false,
  ignoreFonts: true,
});
```

### Separate Style Container

```ts
const contentContainer = document.getElementById("content");
const styleContainer = document.getElementById("styles");

await renderAsync(docxFile, contentContainer, styleContainer);
```

## Supported Features

- ✅ Text and formatting
- ✅ Tables
- ✅ Images
- ✅ Headers and footers
- ✅ Footnotes and endnotes
- ✅ Page breaks
- ✅ Paragraph styles
- ✅ Fonts and colors
- ✅ Text alignment
- ✅ Indents and spacing

## Limitations

- Page breaks work only when explicitly specified in document
- Table of contents not supported (requires fields not implemented)
- Tracked changes in experimental mode
- Comments in experimental mode

## Requirements

- Node.js >= 18.0.0
- JSZip >= 3.10.0
- Modern browsers with ES2020 support

## License

ISC License

## Support

- [GitHub Issues](https://github.com/Andrew-Marty00shenko/docx2html/issues)
- [npm package](https://www.npmjs.com/package/docx-html-renderer)

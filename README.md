# docx-html-renderer

Lightweight Js library to convert DOCX to HTML. Works in plain browser JavaScript (UMD) and in React/TypeScript projects (ESM + typings). Uses JSZip under the hood.

## Installation

```bash
npm install docx-html-renderer
```

📦 **Available on npm**: [docx-html-renderer](https://www.npmjs.com/package/docx-html-renderer)

## Requirements

- **Node.js**: >= 18.0.0
- **JSZip**: Required for DOCX file parsing
- **Browser**: Modern browsers with ES2020 support

## Quick start (UMD)

```html
<!-- dependency: JSZip -->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<!-- library UMD build -->
<script src="lib/docx-html-renderer.js"></script>
<script>
  const container = document.getElementById('container');
  const docData = /* Blob | ArrayBuffer | Uint8Array */;

  docx.renderAsync(docData, container)
    .then(() => console.log('docx: finished'));
</script>
<body>
  <div id="container"></div>
  ...
  <!-- choose and pass a .docx Blob to renderAsync -->
</body>
```

## Usage in ESM / React + TypeScript

```ts
import { renderAsync, parseAsync, renderDocument, defaultOptions } from "docx-html-renderer";

// Basic usage
const renderDocx = async (file: File) => {
  const container = document.getElementById("output");
  await renderAsync(file, container);
};

// With custom options
const renderWithOptions = async (file: File) => {
  const container = document.getElementById("output");
  await renderAsync(file, container, null, {
    className: "my-docx",
    breakPages: true,
    renderHeaders: true,
    renderFooters: true,
  });
};
```

## API

```ts
// renders document into specified element
renderAsync(
  document: Blob | ArrayBuffer | Uint8Array, // any type supported by JSZip.loadAsync
  bodyContainer: HTMLElement,                // element to render document content
  styleContainer: HTMLElement,               // element to render document styles; if null, bodyContainer is used
  options: {
    className: string = 'docx',              // class name/prefix for default and document style classes
    inWrapper: boolean = true,               // enables rendering of wrapper around document content
    hideWrapperOnPrint: boolean = false,     // disable wrapper styles on print
    ignoreWidth: boolean = false,            // disables rendering width of page
    ignoreHeight: boolean = false,           // disables rendering height of page
    ignoreFonts: boolean = false,            // disables fonts rendering
    breakPages: boolean = true,              // enables page breaking on page breaks
    ignoreLastRenderedPageBreak: boolean = true, // disables page breaking on lastRenderedPageBreak elements
    experimental: boolean = false,           // enables experimental features (tab stops calculation)
    trimXmlDeclaration: boolean = true,      // remove xml declaration from xml documents before parsing
    useBase64URL: boolean = false,           // if true, resources use base64 URL, otherwise URL.createObjectURL
    renderChanges: false,                    // experimental rendering of document changes (insertions/deletions)
    renderHeaders: true,                     // enables headers rendering
    renderFooters: true,                     // enables footers rendering
    renderFootnotes: true,                   // enables footnotes rendering
    renderEndnotes: true,                    // enables endnotes rendering
    renderComments: false,                   // enables experimental comments rendering
    renderAltChunks: true,                   // enables altChunks (html parts) rendering
    debug: boolean = false,                  // enables additional logging
  }
): Promise<WordDocument>

// experimental / internal split API
parseAsync(document, options): Promise<WordDocument>
renderDocument(wordDocument, bodyContainer, styleContainer, options): Promise<void>
```

## Notes

- Thumbnails in the demo are for example only and are not part of the library.
- Table of contents relies on fields; fields are not supported yet.

## Page breaks

The library breaks pages when:

- a manual page break `<w:br w:type="page"/>` is inserted
- an application page break `<w:lastRenderedPageBreak/>` is present (set `ignoreLastRenderedPageBreak=false`)
- paragraph page settings change (e.g. portrait → landscape)

Realtime page breaking is not implemented as it would require recalculations on every insertion.

By default `ignoreLastRenderedPageBreak` is `true`.

## Status

The high‑level `renderAsync` API is stable. Internal parsing/rendering implementation details may change.

## Version

Current version: **0.0.8**

## License

ISC License - see [LICENSE](LICENSE) file for details.

## Author

**Andrew Marty00shenko** - [GitHub](https://github.com/Andrew-Marty00shenko)

## Support

- 📖 **Documentation**: Check the API section below
- 🐛 **Issues**: Report bugs on [GitHub](https://github.com/Andrew-Marty00shenko/docx2html/issues)
- ⭐ **Star**: If this library helps you, consider giving it a star!

## Features

- ✅ **DOCX to HTML conversion** - Convert Microsoft Word documents to semantic HTML
- ✅ **React/TypeScript support** - Full TypeScript definitions and React integration
- ✅ **UMD and ESM builds** - Works in browsers and Node.js environments
- ✅ **Semantic HTML output** - Clean, accessible HTML markup
- ✅ **Page breaks support** - Automatic and manual page break handling
- ✅ **Headers and footers** - Document header/footer rendering
- ✅ **Tables and formatting** - Preserve table structure and text formatting
- ✅ **Font and style rendering** - Maintain document styling and typography

# docx-html-renderer

Lightweight library for DOCX → HTML conversion/rendering. Works in native JS (UMD) and React/TypeScript (ESM + types).

## Goal

Goal of this project is to render/convert DOCX document into HTML document with keeping HTML semantic as much as possible.
That means library is limited by HTML capabilities (for example Google Docs renders \*.docx document on canvas as an image).

## Installation

```bash
npm install docx-html-renderer
```

## Quick Start (UMD)

```html
<!--lib uses jszip-->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="lib/docx-html-renderer.js"></script>
<script>
  var docData = <document Blob>;

  docx.renderAsync(docData, document.getElementById("container"))
      .then(x => console.log("docx: finished"));
</script>
<body>
  ...
  <div id="container"></div>
  ...
</body>
```

## Usage in ESM/React+TS

```ts
import { renderAsync, parseAsync, renderDocument, defaultOptions } from "docx-html-renderer";
```

## API

```ts
// renders document into specified element
renderAsync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        className: string = "docx", //class name/prefix for default and document style classes
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        hideWrapperOnPrint: boolean = false, //disable wrapper styles on print
        ignoreWidth: boolean = false, //disables rendering width of page
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreFonts: boolean = false, //disables fonts rendering
        breakPages: boolean = true, //enables page breaking on page breaks
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        experimental: boolean = false, //enables experimental features (tab stops calculation)
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        renderChanges: false, //enables experimental rendering of document changes (inserions/deletions)
        renderHeaders: true, //enables headers rendering
        renderFooters: true, //enables footers rendering
        renderFootnotes: true, //enables footnotes rendering
        renderEndnotes: true, //enables endnotes rendering
        renderComments: false, //enables experimental comments rendering
        renderAltChunks: true, //enables altChunks (html parts) rendering
        debug: boolean = false, //enables additional logging
    }): Promise<WordDocument>

/// ==== experimental / internal API ===
// this API could be used to modify document before rendering
// renderAsync = parseAsync + renderDocument

// parse document and return internal document object
parseAsync(
    document: Blob | ArrayBuffer | Uint8Array,
    options: Options
): Promise<WordDocument>

// render internal document object into specified container
renderDocument(
    wordDocument: WordDocument,
    bodyContainer: HTMLElement,
    styleContainer: HTMLElement,
    options: Options
): Promise<void>
```

## Thumbnails, TOC and etc.

Thumbnails is added only for example and it's not part of library. Library renders DOCX into HTML, so it can't be efficiently used for thumbnails.

Table of contents is built using the TOC fields and there is no efficient way to get table of contents at this point, since fields is not supported yet (http://officeopenxml.com/WPtableOfContents.php)

## Breaks

Currently library does break pages:

- if user/manual page break `<w:br w:type="page"/>` is inserted - when user insert page break
- if application page break `<w:lastRenderedPageBreak/>` is inserted - could be inserted by editor application like MS word (`ignoreLastRenderedPageBreak` should be set to false)
- if page settings for paragraph is changed - ex: user change settings from portrait to landscape page

Realtime page breaking is not implemented because it's requires re-calculation of sizes on each insertion and that could affect performance a lot.

If page breaking is crutual for you, I would recommend:

- try to insert manual break point as much as you could
- try use editors like MS Word, that inserts `<w:lastRenderedPageBreak/>` break points

NOTE: by default `ignoreLastRenderedPageBreak` is set to `true`. You may need to set it to `false`, to make library break by `<w:lastRenderedPageBreak/>` break points

## Status and stability

So far I can't come up with final approach of parsing documents and final structure of API. Only **renderAsync** function is stable and definition shouldn't be changed in future. Inner implementation of parsing and rendering may be changed at any point of time.

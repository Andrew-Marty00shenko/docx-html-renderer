import type { OutputType } from "jszip";

import type { DocumentParser } from "../document-parser";
import type { Relationship } from "../common";
import type { Part } from "../common";
import { FontTablePart } from "../font-table";
import { OpenXmlPackage } from "../common";
import { DocumentPart } from "../document";
import { blobToBase64, resolvePath, splitPath } from "../utils";
import { NumberingPart } from "../numbering";
import { StylesPart } from "../styles";
import { FooterPart, HeaderPart } from "../header-footer";
import { ExtendedPropsPart } from "../document-props";
import { CorePropsPart } from "../document-props";
import { ThemePart } from "../theme";
import { EndnotesPart, FootnotesPart } from "../notes";
import { SettingsPart } from "../settings";
import { CustomPropsPart } from "../document-props";
import { CommentsPart } from "../comments";
import { CommentsExtendedPart } from "../comments";

interface WordDocumentOptions {
  trimXmlDeclaration: boolean;
  keepOrigin: boolean;
  useBase64URL?: boolean;
}

const topLevelRels = [
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    target: "word/document.xml",
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    target: "docProps/app.xml",
  },
  {
    type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    target: "docProps/core.xml",
  },
  {
    type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties",
    target: "docProps/custom.xml",
  },
];

export class WordDocument {
  private _package: OpenXmlPackage;
  private _parser: DocumentParser;
  private _options: WordDocumentOptions;

  rels: Relationship[];
  parts: Part[] = [];
  partsMap: Record<string, Part> = {};

  documentPart: DocumentPart;
  fontTablePart: FontTablePart;
  numberingPart: NumberingPart;
  stylesPart: StylesPart;
  footnotesPart: FootnotesPart;
  endnotesPart: EndnotesPart;
  themePart: ThemePart;
  corePropsPart: CorePropsPart;
  extendedPropsPart: ExtendedPropsPart;
  settingsPart: SettingsPart;
  commentsPart: CommentsPart;
  commentsExtendedPart: CommentsExtendedPart;

  static async load(
    blob: Blob | ArrayBuffer,
    parser: DocumentParser,
    options: WordDocumentOptions,
  ): Promise<WordDocument> {
    const d = new WordDocument();

    d._options = options;
    d._parser = parser;
    d._package = await OpenXmlPackage.load(blob, options);
    d.rels = await d._package.loadRelationships();

    await Promise.all(
      topLevelRels.map((rel) => {
        const r = d.rels.find((x) => x.type === rel.type) ?? rel; //fallback
        return d.loadRelationshipPart(r.target, r.type);
      }),
    );

    return d;
  }

  save(
    type: "blob" | "string" | "uint8array" | "arraybuffer" = "blob",
  ): Promise<Blob | ArrayBuffer | string> {
    return this._package.save(type) as Promise<Blob | ArrayBuffer | string>;
  }

  private async loadRelationshipPart(path: string, type: string): Promise<Part> {
    if (this.partsMap[path]) return this.partsMap[path];

    if (!this._package.get(path)) return null;

    let part: Part = null;

    switch (type) {
      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
        this.documentPart = part = new DocumentPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
        this.fontTablePart = part = new FontTablePart(this._package, path);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
        this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
        this.stylesPart = part = new StylesPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
        this.themePart = part = new ThemePart(this._package, path);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
        this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
        this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer":
        part = new FooterPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header":
        part = new HeaderPart(this._package, path, this._parser);
        break;

      case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties":
        this.corePropsPart = part = new CorePropsPart(this._package, path);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties":
        this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
        break;

      case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties":
        part = new CustomPropsPart(this._package, path);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings":
        this.settingsPart = part = new SettingsPart(this._package, path);
        break;

      case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
        this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
        break;

      case "http://schemas.microsoft.com/office/2011/relationships/commentsExtended":
        this.commentsExtendedPart = part = new CommentsExtendedPart(this._package, path);
        break;
    }

    if (part == null) return Promise.resolve(null);

    this.partsMap[path] = part;
    this.parts.push(part);

    await part.load();

    if (part.rels?.length > 0) {
      const [folder] = splitPath(part.path);
      await Promise.all(
        part.rels.map((rel) =>
          this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type),
        ),
      );
    }

    return part;
  }

  async loadDocumentImage(id: string, part?: Part): Promise<string> {
    const x = await this.loadResource(part ?? this.documentPart, id, "blob");
    return this.blobToURL(x);
  }

  async loadNumberingImage(id: string): Promise<string> {
    const x = await this.loadResource(this.numberingPart, id, "blob");
    return this.blobToURL(x);
  }

  async loadFont(id: string, key: string): Promise<string> {
    const x = await this.loadResource(this.fontTablePart, id, "uint8array");
    return x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x;
  }

  async loadAltChunk(id: string, part?: Part): Promise<string> {
    return await this.loadResource(part ?? this.documentPart, id, "string");
  }

  private blobToURL(blob: Blob): string | Promise<string> {
    if (!blob) return null;

    if (this._options.useBase64URL) {
      return blobToBase64(blob);
    }

    return URL.createObjectURL(blob);
  }

  findPartByRelId(id: string, basePart: Part = null) {
    const rel = (basePart.rels ?? this.rels).find((r) => r.id == id);
    const folder = basePart ? splitPath(basePart.path)[0] : "";
    return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
  }

  getPathById(part: Part, id: string): string {
    const rel = part.rels.find((x) => x.id == id);
    const [folder] = splitPath(part.path);
    return rel ? resolvePath(rel.target, folder) : null;
  }

  private loadResource(part: Part, id: string, outputType: OutputType) {
    const path = this.getPathById(part, id);
    return path ? this._package.load(path, outputType) : Promise.resolve(null);
  }
}

function deobfuscate(data: Uint8Array, guidKey: string): Uint8Array {
  const len = 16;
  const trimmed = guidKey.replace(/{|}|-/g, "");
  const numbers = new Array(len);

  for (let i = 0; i < len; i++) numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);

  for (let i = 0; i < 32; i++) data[i] = data[i] ^ numbers[i % len];

  return data;
}

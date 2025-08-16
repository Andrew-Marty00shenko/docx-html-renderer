import JSZip from "jszip";
import { parseXmlString, XmlParser } from "../parser";
import { splitPath } from "../utils";
import type { Relationship } from "./relationship";
import { parseRelationships } from "./relationship";

interface OpenXmlPackageOptions {
  trimXmlDeclaration: boolean;
  keepOrigin: boolean;
}

export class OpenXmlPackage {
  xmlParser: XmlParser = new XmlParser();

  constructor(
    private _zip: JSZip,
    public options: OpenXmlPackageOptions,
  ) {}

  get(path: string): Nullable<JSZip.JSZipObject> {
    const p = normalizePath(path);
    return this._zip.files[p] ?? this._zip.files[p.replace(/\//g, "\\")] ?? null;
  }

  update(path: string, content: string | Uint8Array | Blob) {
    this._zip.file(path, content);
  }

  static async load(
    input: Blob | ArrayBuffer | Uint8Array,
    options: OpenXmlPackageOptions,
  ): Promise<OpenXmlPackage> {
    const zip = await JSZip.loadAsync(input);
    return new OpenXmlPackage(zip, options);
  }

  save(type: JSZip.OutputType = "blob"): Promise<unknown> {
    return this._zip.generateAsync({ type });
  }

  load(
    path: string,
    type: JSZip.OutputType = "string",
  ): Promise<Nullable<string | Uint8Array | Blob>> {
    return this.get(path)?.async(type) ?? Promise.resolve(null);
  }

  async loadRelationships(path: Nullable<string> = null): Promise<Nullable<Relationship[]>> {
    let relsPath = `_rels/.rels`;

    if (path != null) {
      const [f, fn] = splitPath(path);
      relsPath = `${f}_rels/${fn}.rels`;
    }

    const txt = await this.load(relsPath);
    return txt
      ? parseRelationships(this.parseXmlDocument(txt as string).firstElementChild, this.xmlParser)
      : null;
  }

  /** @internal */
  parseXmlDocument(txt: string): Document {
    return parseXmlString(txt, this.options.trimXmlDeclaration);
  }
}

function normalizePath(path: string) {
  return path.startsWith("/") ? path.substr(1) : path;
}

import type { OpenXmlPackage } from "../common";
import { Part } from "../common";
import type { DocumentParser } from "../document-parser";
import type { DocumentElement } from "./document";

export class DocumentPart extends Part {
  private _documentParser: DocumentParser;

  constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
    super(pkg, path);
    this._documentParser = parser;
  }

  body: DocumentElement;

  parseXml(root: Element) {
    this.body = this._documentParser.parseDocumentFile(root);
  }
}

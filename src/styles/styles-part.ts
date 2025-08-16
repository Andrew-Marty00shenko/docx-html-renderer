import type { OpenXmlPackage } from "../common";
import { Part } from "../common";
import type { DocumentParser } from "../document-parser";
import type { IDomStyle } from "../document";

export class StylesPart extends Part {
  styles: IDomStyle[];

  private _documentParser: DocumentParser;

  constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
    super(pkg, path);
    this._documentParser = parser;
  }

  parseXml(root: Element) {
    this.styles = this._documentParser.parseStylesFile(root);
  }
}

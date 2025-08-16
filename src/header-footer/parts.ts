import type { OpenXmlPackage } from "../common";
import { Part } from "../common";
import type { DocumentParser } from "../document-parser";
import type { OpenXmlElement } from "../document";
import { WmlHeader, WmlFooter } from "./elements";

export abstract class BaseHeaderFooterPart<T extends OpenXmlElement = OpenXmlElement> extends Part {
  rootElement: T;

  private _documentParser: DocumentParser;

  constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
    super(pkg, path);
    this._documentParser = parser;
  }

  parseXml(root: Element) {
    this.rootElement = this.createRootElement();
    this.rootElement.children = this._documentParser.parseBodyElements(root);
  }

  protected abstract createRootElement(): T;
}

export class HeaderPart extends BaseHeaderFooterPart<WmlHeader> {
  protected createRootElement(): WmlHeader {
    return new WmlHeader();
  }
}

export class FooterPart extends BaseHeaderFooterPart<WmlFooter> {
  protected createRootElement(): WmlFooter {
    return new WmlFooter();
  }
}

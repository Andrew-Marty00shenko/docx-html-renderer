import { serializeXmlString } from "../parser";
import type { OpenXmlPackage } from "./open-xml-package";
import type { Relationship } from "./relationship";

export class Part {
  protected _xmlDocument: Document;

  rels: Relationship[];

  constructor(
    protected _package: OpenXmlPackage,
    public path: string,
  ) {}

  async load(): Promise<void> {
    this.rels = await this._package.loadRelationships(this.path);

    const xmlText = await this._package.load(this.path);
    if (xmlText && typeof xmlText === "string") {
      const xmlDoc = this._package.parseXmlDocument(xmlText);

      if (this._package.options.keepOrigin) {
        this._xmlDocument = xmlDoc;
      }

      this.parseXml(xmlDoc.firstElementChild);
    }
  }

  save() {
    this._package.update(this.path, serializeXmlString(this._xmlDocument));
  }

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  protected parseXml(_root: Element) {}
}

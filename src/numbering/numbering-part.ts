import type { OpenXmlPackage } from "../common";
import { Part } from "../common";
import type { DocumentParser } from "../document-parser";
import type { IDomNumbering } from "../document";
import type {
  AbstractNumbering,
  Numbering,
  NumberingBulletPicture,
  NumberingPartProperties,
} from "./numbering";
import { parseNumberingPart } from "./numbering";

export class NumberingPart extends Part implements NumberingPartProperties {
  private _documentParser: DocumentParser;

  constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
    super(pkg, path);
    this._documentParser = parser;
  }

  numberings: Numbering[];
  abstractNumberings: AbstractNumbering[];
  bulletPictures: NumberingBulletPicture[];

  domNumberings: IDomNumbering[];

  parseXml(root: Element) {
    Object.assign(this, parseNumberingPart(root, this._package.xmlParser));
    this.domNumberings = this._documentParser.parseNumberingFile(root);
  }
}

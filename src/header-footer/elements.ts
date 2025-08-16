import type { DomType } from "../document";
import { OpenXmlElementBase } from "../document";

export class WmlHeader extends OpenXmlElementBase {
  type: DomType = "header";
}

export class WmlFooter extends OpenXmlElementBase {
  type: DomType = "footer";
}

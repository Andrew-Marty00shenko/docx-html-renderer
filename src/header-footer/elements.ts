import { OpenXmlElementBase, DomType } from "../document";

export class WmlHeader extends OpenXmlElementBase {
  type: DomType = DomType.Header;
}

export class WmlFooter extends OpenXmlElementBase {
  type: DomType = DomType.Footer;
}

import type { OpenXmlElementBase, OpenXmlElement, DomType } from "../document";

export abstract class WmlBaseNote implements OpenXmlElementBase {
  type: DomType;
  id: string;
  noteType: string;
  children: OpenXmlElement[] = [];
}

export class WmlFootnote extends WmlBaseNote {
  type: DomType = "footnote";
}

export class WmlEndnote extends WmlBaseNote {
  type: DomType = "endnote";
}

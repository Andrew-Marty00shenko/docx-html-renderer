import type { DomType } from "../document";
import { OpenXmlElementBase } from "../document";

export class WmlComment extends OpenXmlElementBase {
  type: DomType = "comment";
  id: string;
  author: string;
  initials: string;
  date: string;
}

export class WmlCommentReference extends OpenXmlElementBase {
  type: DomType = "commentReference";

  constructor(public id?: string) {
    super();
  }
}

export class WmlCommentRangeStart extends OpenXmlElementBase {
  type: DomType = "commentRangeStart";

  constructor(public id?: string) {
    super();
  }
}
export class WmlCommentRangeEnd extends OpenXmlElementBase {
  type: DomType = "commentRangeEnd";

  constructor(public id?: string) {
    super();
  }
}

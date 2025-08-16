import type { XmlParser } from "../parser";
import type { OpenXmlElement } from "./dom";

export interface WmlBookmarkStart extends OpenXmlElement {
  id: string;
  name: string;
  colFirst: number;
  colLast: number;
}

interface WmlBookmarkEnd extends OpenXmlElement {
  id: string;
}

export function parseBookmarkStart(elem: Element, xml: XmlParser): WmlBookmarkStart {
  return {
    type: "bookmarkStart",
    id: xml.attr(elem, "id"),
    name: xml.attr(elem, "name"),
    colFirst: xml.intAttr(elem, "colFirst"),
    colLast: xml.intAttr(elem, "colLast"),
  };
}

export function parseBookmarkEnd(elem: Element, xml: XmlParser): WmlBookmarkEnd {
  return {
    type: "bookmarkEnd",
    id: xml.attr(elem, "id"),
  };
}

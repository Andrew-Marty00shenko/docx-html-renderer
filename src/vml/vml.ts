import type { DocumentParser } from "../document-parser";
import { OpenXmlElementBase, DomType } from "../document/dom";
import xml from "../parser/xml-parser";

export class VmlElement extends OpenXmlElementBase {
  type: DomType = DomType.VmlElement;
  tagName: string;
  cssStyleText?: string;
  attrs: Record<string, string> = {};
  imageHref?: {
    id: string;
    title: string;
  };
}

export function parseVmlElement(elem: Element, parser: DocumentParser): VmlElement {
  const result = new VmlElement();

  switch (elem.localName) {
    case "rect":
      result.tagName = "rect";
      result.attrs = { width: "100%", height: "100%" };
      break;

    case "oval":
      result.tagName = "ellipse";
      result.attrs = { cx: "50%", cy: "50%", rx: "50%", ry: "50%" };
      break;

    case "line":
      result.tagName = "line";
      break;

    case "shape":
      result.tagName = "g";
      break;

    case "textbox":
      result.tagName = "foreignObject";
      result.attrs = { width: "100%", height: "100%" };
      break;

    default:
      return null;
  }

  for (const at of xml.attrs(elem)) {
    switch (at.localName) {
      case "style":
        result.cssStyleText = at.value;
        break;

      case "fillcolor":
        result.attrs.fill = at.value;
        break;

      case "from":
        const [x1, y1] = at.value.split(",");
        result.attrs.x1 = x1;
        result.attrs.y1 = y1;
        break;

      case "to":
        const [x2, y2] = at.value.split(",");
        result.attrs.x2 = x2;
        result.attrs.y2 = y2;
        break;
    }
  }

  for (const el of xml.elements(elem)) {
    switch (el.localName) {
      case "stroke":
        result.attrs.stroke = xml.attr(el, "color");
        result.attrs["stroke-width"] =
          xml.lengthAttr(el, "weight", { mul: 1 / 12700, unit: "" }) ?? "1px";
        break;

      case "imagedata":
        result.tagName = "image";
        result.attrs = { width: "100%", height: "100%" };
        result.imageHref = {
          id: xml.attr(el, "id"),
          title: xml.attr(el, "title"),
        };
        break;

      case "txbxContent":
        result.children.push(...parser.parseBodyElements(el));
        break;

      default:
        const child = parseVmlElement(el, parser);
        if (child) {
          result.children.push(child);
        }
        break;
    }
  }

  return result;
}

import type { OpenXmlElement } from "./dom";
import type { CommonProperties, Length } from "./common";
import { ns, parseCommonProperty } from "./common";
import type { Borders } from "./border";
import type { SectionProperties } from "./section";
import { parseSectionProperties } from "./section";
import type { LineSpacing } from "./line-spacing";
import { parseLineSpacing } from "./line-spacing";
import type { XmlParser } from "../parser";
import type { RunProperties } from "./run";
import { parseRunProperties } from "./run";

export interface WmlParagraph extends OpenXmlElement, ParagraphProperties {}

export interface ParagraphProperties extends CommonProperties {
  sectionProps: SectionProperties;
  tabs: ParagraphTab[];
  numbering: ParagraphNumbering;

  border: Borders;
  textAlignment: "auto" | "baseline" | "bottom" | "center" | "top" | string;
  lineSpacing: LineSpacing;
  keepLines: boolean;
  keepNext: boolean;
  pageBreakBefore: boolean;
  outlineLevel: number;
  styleName?: string;

  runProps: RunProperties;
}

export interface ParagraphTab {
  style: "bar" | "center" | "clear" | "decimal" | "end" | "num" | "start" | "left" | "right";
  leader: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
  position: Length;
}

export interface ParagraphNumbering {
  id: string;
  level: number;
}

export function parseParagraphProperties(elem: Element, xml: XmlParser): ParagraphProperties {
  const result = {} as ParagraphProperties;

  for (const el of xml.elements(elem)) {
    parseParagraphProperty(el, result, xml);
  }

  return result;
}

export function parseParagraphProperty(elem: Element, props: ParagraphProperties, xml: XmlParser) {
  if (elem.namespaceURI != ns.wordml) return false;

  if (parseCommonProperty(elem, props, xml)) return true;

  switch (elem.localName) {
    case "tabs":
      props.tabs = parseTabs(elem, xml);
      break;

    case "sectPr":
      props.sectionProps = parseSectionProperties(elem, xml);
      break;

    case "numPr":
      props.numbering = parseNumbering(elem, xml);
      break;

    case "spacing":
      props.lineSpacing = parseLineSpacing(elem, xml);
      return false; // TODO
      break;

    case "textAlignment":
      props.textAlignment = xml.attr(elem, "val");
      return false; //TODO
      break;

    case "keepLines":
      props.keepLines = xml.boolAttr(elem, "val", true);
      break;

    case "keepNext":
      props.keepNext = xml.boolAttr(elem, "val", true);
      break;

    case "pageBreakBefore":
      props.pageBreakBefore = xml.boolAttr(elem, "val", true);
      break;

    case "outlineLvl":
      props.outlineLevel = xml.intAttr(elem, "val");
      break;

    case "pStyle":
      props.styleName = xml.attr(elem, "val");
      break;

    case "rPr":
      props.runProps = parseRunProperties(elem, xml);
      break;

    default:
      return false;
  }

  return true;
}

export function parseTabs(elem: Element, xml: XmlParser): ParagraphTab[] {
  return xml.elements(elem, "tab").map(
    (e) =>
      ({
        position: xml.lengthAttr(e, "pos"),
        leader: xml.attr(e, "leader"),
        style: xml.attr(e, "val"),
      }) as ParagraphTab,
  );
}

export function parseNumbering(elem: Element, xml: XmlParser): ParagraphNumbering {
  const result = {} as ParagraphNumbering;

  for (const e of xml.elements(elem)) {
    switch (e.localName) {
      case "numId":
        result.id = xml.attr(e, "val");
        break;

      case "ilvl":
        result.level = xml.intAttr(e, "val");
        break;
    }
  }

  return result;
}

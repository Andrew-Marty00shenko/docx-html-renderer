import type { XmlParser } from "../parser/xml-parser";
import type { CommonProperties} from "./common";
import { parseCommonProperty } from "./common";
import type { OpenXmlElement } from "./dom";

export interface WmlRun extends OpenXmlElement, RunProperties {
  id?: string;
  verticalAlign?: string;
  fieldRun?: boolean;
}

export interface RunProperties extends CommonProperties {}

export function parseRunProperties(elem: Element, xml: XmlParser): RunProperties {
  const result = {} as RunProperties;

  for (const el of xml.elements(elem)) {
    parseRunProperty(el, result, xml);
  }

  return result;
}

export function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser) {
  if (parseCommonProperty(elem, props, xml)) return true;

  return false;
}

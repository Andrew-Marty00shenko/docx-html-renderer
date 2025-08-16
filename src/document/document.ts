import type { OpenXmlElement } from "./dom";
import type { SectionProperties } from "./section";

export interface DocumentElement extends OpenXmlElement {
  props: SectionProperties & Record<string, unknown>;
}

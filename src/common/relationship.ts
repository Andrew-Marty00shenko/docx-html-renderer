import type { XmlParser } from "../parser";

export interface Relationship {
  id: string;
  type: RelationshipTypes | string;
  target: string;
  targetMode: "" | "External" | string;
}

type RelationshipTypes =
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  | "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
  | "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/custom-properties"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  | "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
  | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

export function parseRelationships(root: Element, xml: XmlParser): Relationship[] {
  return xml.elements(root).map(
    (e) =>
      ({
        id: xml.attr(e, "Id"),
        type: xml.attr(e, "Type"),
        target: xml.attr(e, "Target"),
        targetMode: xml.attr(e, "TargetMode"),
      }) as Relationship,
  );
}

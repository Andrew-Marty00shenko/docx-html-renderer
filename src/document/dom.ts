export type DomType =
  | "document"
  | "paragraph"
  | "run"
  | "break"
  | "noBreakHyphen"
  | "table"
  | "row"
  | "cell"
  | "hyperlink"
  | "smartTag"
  | "drawing"
  | "image"
  | "text"
  | "tab"
  | "symbol"
  | "bookmarkStart"
  | "bookmarkEnd"
  | "footer"
  | "header"
  | "footnoteReference"
  | "endnoteReference"
  | "footnote"
  | "endnote"
  | "simpleField"
  | "complexField"
  | "instruction"
  | "vmlPicture"
  | "mmlMath"
  | "mmlMathParagraph"
  | "mmlFraction"
  | "mmlFunction"
  | "mmlFunctionName"
  | "mmlNumerator"
  | "mmlDenominator"
  | "mmlRadical"
  | "mmlBase"
  | "mmlDegree"
  | "mmlSuperscript"
  | "mmlSubscript"
  | "mmlPreSubSuper"
  | "mmlSubArgument"
  | "mmlSuperArgument"
  | "mmlNary"
  | "mmlDelimiter"
  | "mmlRun"
  | "mmlEquationArray"
  | "mmlLimit"
  | "mmlLimitLower"
  | "mmlMatrix"
  | "mmlMatrixRow"
  | "mmlBox"
  | "mmlBar"
  | "mmlGroupChar"
  | "vmlElement"
  | "inserted"
  | "deleted"
  | "deletedText"
  | "comment"
  | "commentReference"
  | "commentRangeStart"
  | "commentRangeEnd"
  | "altChunk";

export interface OpenXmlElement {
  type: DomType;
  children?: OpenXmlElement[];
  cssStyle?: Record<string, string>;
  props?: Record<string, unknown>;

  styleName?: string; //style name
  className?: string; //class mods

  parent?: OpenXmlElement;
}

export abstract class OpenXmlElementBase implements OpenXmlElement {
  type: DomType;
  children?: OpenXmlElement[] = [];
  cssStyle?: Record<string, string> = {};
  props?: Record<string, unknown>;

  className?: string;
  styleName?: string;

  parent?: OpenXmlElement;
}

export interface WmlHyperlink extends OpenXmlElement {
  id?: string;
  anchor?: string;
}

export interface WmlAltChunk extends OpenXmlElement {
  id?: string;
}

export interface WmlSmartTag extends OpenXmlElement {
  uri?: string;
  element?: string;
}

export interface WmlNoteReference extends OpenXmlElement {
  id: string;
}

export interface WmlBreak extends OpenXmlElement {
  break: "page" | "lastRenderedPageBreak" | "textWrapping";
}

export interface WmlText extends OpenXmlElement {
  text: string;
}

export interface WmlSymbol extends OpenXmlElement {
  font: string;
  char: string;
}

export interface WmlTable extends OpenXmlElement {
  columns?: WmlTableColumn[];
  cellStyle?: Record<string, string>;

  colBandSize?: number;
  rowBandSize?: number;
}

export interface WmlTableRow extends OpenXmlElement {
  isHeader?: boolean;
  gridBefore?: number;
  gridAfter?: number;
}

export interface WmlTableCell extends OpenXmlElement {
  verticalMerge?: "restart" | "continue" | string;
  span?: number;
}

export interface IDomImage extends OpenXmlElement {
  src: string;
}

export interface WmlTableColumn {
  width?: string;
}

export interface IDomNumbering {
  id: string;
  level: number;
  start: number;
  pStyleName: string;
  pStyle: Record<string, string>;
  rStyle: Record<string, string>;
  levelText?: string;
  suff: string;
  format?: string;
  bullet?: NumberingPicBullet;
}

export interface NumberingPicBullet {
  id: number;
  src: string;
  style?: string;
}

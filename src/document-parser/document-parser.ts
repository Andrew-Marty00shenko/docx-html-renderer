import type {
  WmlTable,
  IDomNumbering,
  WmlHyperlink,
  WmlSmartTag,
  IDomImage,
  OpenXmlElement,
  WmlTableColumn,
  WmlTableCell,
  WmlTableRow,
  NumberingPicBullet,
  WmlText,
  WmlSymbol,
  WmlBreak,
  WmlNoteReference,
  WmlAltChunk,
  WmlParagraph,
  SectionProperties,
  DocumentElement,
  WmlRun,
  IDomStyle,
  IDomSubStyle,
  WmlFieldChar,
  WmlFieldSimple,
  WmlInstructionText,
  LengthUsageType,
  DomType,
} from "../document";
import {
  parseParagraphProperties,
  parseParagraphProperty,
  parseSectionProperties,
  parseRunProperties,
  parseBookmarkEnd,
  parseBookmarkStart,
  convertLength,
  LengthUsage,
} from "../document";
import { XmlParser } from "../parser";
import { parseVmlElement } from "../vml";
import {
  WmlComment,
  WmlCommentRangeEnd,
  WmlCommentRangeStart,
  WmlCommentReference,
} from "../comments";
import { encloseFontFamily } from "../utils";

export var autos = {
  shd: "inherit",
  color: "black",
  borderColor: "black",
  highlight: "transparent",
};

const supportedNamespaceURIs = [];

const mmlTagMap: Record<string, DomType> = {
  oMath: "mmlMath",
  oMathPara: "mmlMathParagraph",
  f: "mmlFraction",
  func: "mmlFunction",
  fName: "mmlFunctionName",
  num: "mmlNumerator",
  den: "mmlDenominator",
  rad: "mmlRadical",
  deg: "mmlDegree",
  e: "mmlBase",
  sSup: "mmlSuperscript",
  sSub: "mmlSubscript",
  sPre: "mmlPreSubSuper",
  sup: "mmlSuperArgument",
  sub: "mmlSubArgument",
  d: "mmlDelimiter",
  nary: "mmlNary",
  eqArr: "mmlEquationArray",
  lim: "mmlLimit",
  limLow: "mmlLimitLower",
  m: "mmlMatrix",
  mr: "mmlMatrixRow",
  box: "mmlBox",
  bar: "mmlBar",
  groupChr: "mmlGroupChar",
};

export interface DocumentParserOptions {
  ignoreWidth: boolean;
  debug: boolean;
}

export class DocumentParser {
  options: DocumentParserOptions;
  public xmlParser: XmlParser;

  constructor(options?: Partial<DocumentParserOptions>) {
    this.options = {
      ignoreWidth: false,
      debug: false,
      ...options,
    };
    this.xmlParser = new XmlParser();
  }

  parseNotes(
    xmlDoc: Element,
    elemName: string,
    elemClass: new () => {
      id: string;
      noteType: string;
      children: OpenXmlElement[];
      type: DomType;
    },
  ): { id: string; noteType: string; children: OpenXmlElement[]; type: DomType }[] {
    const result = [];

    for (const el of this.xmlParser.elements(xmlDoc, elemName)) {
      const node = new elemClass();
      node.id = this.xmlParser.attr(el, "id");
      node.noteType = this.xmlParser.attr(el, "type");
      node.children = this.parseBodyElements(el);
      result.push(node);
    }

    return result;
  }

  parseComments(xmlDoc: Element): WmlComment[] {
    const result = [];

    for (const el of this.xmlParser.elements(xmlDoc, "comment")) {
      const item = new WmlComment();
      item.id = this.xmlParser.attr(el, "id");
      item.author = this.xmlParser.attr(el, "author");
      item.initials = this.xmlParser.attr(el, "initials");
      item.date = this.xmlParser.attr(el, "date");
      item.children = this.parseBodyElements(el);
      result.push(item);
    }

    return result;
  }

  parseDocumentFile(xmlDoc: Element): DocumentElement {
    const xbody = this.xmlParser.element(xmlDoc, "body");
    const background = this.xmlParser.element(xmlDoc, "background");
    const sectPr = this.xmlParser.element(xbody, "sectPr");

    return {
      type: "document",
      children: this.parseBodyElements(xbody),
      props: sectPr
        ? (parseSectionProperties(sectPr, this.xmlParser) as SectionProperties &
            Record<string, unknown>)
        : ({} as SectionProperties & Record<string, unknown>),
      cssStyle: background ? this.parseBackground(background) : {},
    };
  }

  parseBackground(elem: Element): Record<string, string> {
    const result = {};
    const color = xmlUtil.colorAttr(elem, "color");

    if (color) {
      result["background-color"] = color;
    }

    return result;
  }

  parseBodyElements(element: Element): OpenXmlElement[] {
    const children = [];

    for (const elem of this.xmlParser.elements(element)) {
      switch (elem.localName) {
        case "p":
          children.push(this.parseParagraph(elem));
          break;

        case "altChunk":
          children.push(this.parseAltChunk(elem));
          break;

        case "tbl":
          children.push(this.parseTable(elem));
          break;

        case "sdt":
          children.push(...this.parseSdt(elem, (e) => this.parseBodyElements(e)));
          break;
      }
    }

    return children;
  }

  parseStylesFile(xstyles: Element): IDomStyle[] {
    const result = [];

    xmlUtil.foreach(xstyles, (n) => {
      switch (n.localName) {
        case "style":
          result.push(this.parseStyle(n));
          break;

        case "docDefaults":
          result.push(this.parseDefaultStyles(n));
          break;
      }
    });

    return result;
  }

  parseDefaultStyles(node: Element): IDomStyle {
    const result = {
      id: null,
      name: null,
      target: null,
      basedOn: null,
      styles: [],
    } as IDomStyle;

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "rPrDefault":
          var rPr = this.xmlParser.element(c, "rPr");

          if (rPr)
            result.styles.push({
              target: "span",
              values: this.parseDefaultProperties(rPr, {}),
            });
          break;

        case "pPrDefault":
          var pPr = this.xmlParser.element(c, "pPr");

          if (pPr)
            result.styles.push({
              target: "p",
              values: this.parseDefaultProperties(pPr, {}),
            });
          break;
      }
    });

    return result;
  }

  parseStyle(node: Element): IDomStyle {
    const result = {
      id: this.xmlParser.attr(node, "styleId"),
      isDefault: this.xmlParser.boolAttr(node, "default"),
      name: null,
      target: null,
      basedOn: null,
      styles: [],
      linked: null,
    } as IDomStyle;

    switch (this.xmlParser.attr(node, "type")) {
      case "paragraph":
        result.target = "p";
        break;
      case "table":
        result.target = "table";
        break;
      case "character":
        result.target = "span";
        break;
      //case "numbering": result.target = "p"; break;
    }

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case "basedOn":
          result.basedOn = this.xmlParser.attr(n, "val");
          break;

        case "name":
          result.name = this.xmlParser.attr(n, "val");
          break;

        case "link":
          result.linked = this.xmlParser.attr(n, "val");
          break;

        case "next":
          result.next = this.xmlParser.attr(n, "val");
          break;

        case "aliases":
          result.aliases = this.xmlParser.attr(n, "val").split(",");
          break;

        case "pPr":
          result.styles.push({
            target: "p",
            values: this.parseDefaultProperties(n, {}),
          });
          result.paragraphProps = parseParagraphProperties(n, this.xmlParser);
          break;

        case "rPr":
          result.styles.push({
            target: "span",
            values: this.parseDefaultProperties(n, {}),
          });
          result.runProps = parseRunProperties(n, this.xmlParser);
          break;

        case "tblPr":
        case "tcPr":
          result.styles.push({
            target: "td", //TODO: maybe move to processor
            values: this.parseDefaultProperties(n, {}),
          });
          break;

        case "tblStylePr":
          for (const s of this.parseTableStyle(n)) result.styles.push(s);
          break;

        case "rsid":
        case "qFormat":
        case "hidden":
        case "semiHidden":
        case "unhideWhenUsed":
        case "autoRedefine":
        case "uiPriority":
          //TODO: ignore
          break;

        default:
          if (this.options.debug) {
            // eslint-disable-next-line no-console
            console.warn(`DOCX: Unknown style element: ${n.localName}`);
          }
      }
    });

    return result;
  }

  parseTableStyle(node: Element): IDomSubStyle[] {
    const result = [];

    const type = this.xmlParser.attr(node, "type");
    let selector = "";
    let modificator = "";

    switch (type) {
      case "firstRow":
        modificator = ".first-row";
        selector = "tr.first-row td";
        break;
      case "lastRow":
        modificator = ".last-row";
        selector = "tr.last-row td";
        break;
      case "firstCol":
        modificator = ".first-col";
        selector = "td.first-col";
        break;
      case "lastCol":
        modificator = ".last-col";
        selector = "td.last-col";
        break;
      case "band1Vert":
        modificator = ":not(.no-vband)";
        selector = "td.odd-col";
        break;
      case "band2Vert":
        modificator = ":not(.no-vband)";
        selector = "td.even-col";
        break;
      case "band1Horz":
        modificator = ":not(.no-hband)";
        selector = "tr.odd-row";
        break;
      case "band2Horz":
        modificator = ":not(.no-hband)";
        selector = "tr.even-row";
        break;
      default:
        return [];
    }

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case "pPr":
          result.push({
            target: `${selector} p`,
            mod: modificator,
            values: this.parseDefaultProperties(n, {}),
          });
          break;

        case "rPr":
          result.push({
            target: `${selector} span`,
            mod: modificator,
            values: this.parseDefaultProperties(n, {}),
          });
          break;

        case "tblPr":
        case "tcPr":
          result.push({
            target: selector, //TODO: maybe move to processor
            mod: modificator,
            values: this.parseDefaultProperties(n, {}),
          });
          break;
      }
    });

    return result;
  }

  parseNumberingFile(xnums: Element): IDomNumbering[] {
    const result = [];
    const mapping = {};
    const bullets = [];

    xmlUtil.foreach(xnums, (n) => {
      switch (n.localName) {
        case "abstractNum":
          this.parseAbstractNumbering(n, bullets).forEach((x) => result.push(x));
          break;

        case "numPicBullet":
          bullets.push(this.parseNumberingPicBullet(n));
          break;

        case "num":
          var numId = this.xmlParser.attr(n, "numId");
          var abstractNumId = this.xmlParser.elementAttr(n, "abstractNumId", "val");
          mapping[abstractNumId] = numId;
          break;
      }
    });

    result.forEach((x) => (x.id = mapping[x.id]));

    return result;
  }

  parseNumberingPicBullet(elem: Element): NumberingPicBullet {
    const pict = this.xmlParser.element(elem, "pict");
    const shape = pict && this.xmlParser.element(pict, "shape");
    const imagedata = shape && this.xmlParser.element(shape, "imagedata");

    return imagedata
      ? {
          id: this.xmlParser.intAttr(elem, "numPicBulletId"),
          src: this.xmlParser.attr(imagedata, "id"),
          style: this.xmlParser.attr(shape, "style"),
        }
      : null;
  }

  parseAbstractNumbering(node: Element, bullets: NumberingPicBullet[]): IDomNumbering[] {
    const result = [];
    const id = this.xmlParser.attr(node, "abstractNumId");

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case "lvl":
          result.push(this.parseNumberingLevel(id, n, bullets));
          break;
      }
    });

    return result;
  }

  parseNumberingLevel(id: string, node: Element, bullets: NumberingPicBullet[]): IDomNumbering {
    const result: IDomNumbering = {
      id: id,
      level: this.xmlParser.intAttr(node, "ilvl"),
      start: 1,
      pStyleName: undefined,
      pStyle: {},
      rStyle: {},
      suff: "tab",
    };

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case "start":
          result.start = this.xmlParser.intAttr(n, "val");
          break;

        case "pPr":
          this.parseDefaultProperties(n, result.pStyle);
          break;

        case "rPr":
          this.parseDefaultProperties(n, result.rStyle);
          break;

        case "lvlPicBulletId":
          var id = this.xmlParser.intAttr(n, "val");
          result.bullet = bullets.find((x) => x?.id == id);
          break;

        case "lvlText":
          result.levelText = this.xmlParser.attr(n, "val");
          break;

        case "pStyle":
          result.pStyleName = this.xmlParser.attr(n, "val");
          break;

        case "numFmt":
          result.format = this.xmlParser.attr(n, "val");
          break;

        case "suff":
          result.suff = this.xmlParser.attr(n, "val");
          break;
      }
    });

    return result;
  }

  parseSdt(node: Element, parser: (elem: Element) => OpenXmlElement[]): OpenXmlElement[] {
    const sdtContent = this.xmlParser.element(node, "sdtContent");
    return sdtContent ? parser(sdtContent) : [];
  }

  parseInserted(node: Element, parentParser: (elem: Element) => OpenXmlElement): OpenXmlElement {
    return {
      type: "inserted",
      children: parentParser(node)?.children ?? [],
    } as OpenXmlElement;
  }

  parseDeleted(node: Element, parentParser: (elem: Element) => OpenXmlElement): OpenXmlElement {
    return {
      type: "deleted",
      children: parentParser(node)?.children ?? [],
    } as OpenXmlElement;
  }

  parseAltChunk(node: Element): WmlAltChunk {
    return { type: "altChunk", children: [], id: this.xmlParser.attr(node, "id") };
  }

  parseParagraph(node: Element): OpenXmlElement {
    const result = { type: "paragraph", children: [] } as WmlParagraph;

    for (const el of this.xmlParser.elements(node)) {
      switch (el.localName) {
        case "pPr":
          this.parseParagraphProperties(el, result);
          break;

        case "r":
          result.children.push(this.parseRun(el, result));
          break;

        case "hyperlink":
          result.children.push(this.parseHyperlink(el, result));
          break;

        case "smartTag":
          result.children.push(this.parseSmartTag(el, result));
          break;

        case "bookmarkStart":
          result.children.push(parseBookmarkStart(el, this.xmlParser));
          break;

        case "bookmarkEnd":
          result.children.push(parseBookmarkEnd(el, this.xmlParser));
          break;

        case "commentRangeStart":
          result.children.push(new WmlCommentRangeStart(this.xmlParser.attr(el, "id")));
          break;

        case "commentRangeEnd":
          result.children.push(new WmlCommentRangeEnd(this.xmlParser.attr(el, "id")));
          break;

        case "oMath":
        case "oMathPara":
          result.children.push(this.parseMathElement(el));
          break;

        case "sdt":
          result.children.push(...this.parseSdt(el, (e) => this.parseParagraph(e).children));
          break;

        case "ins":
          result.children.push(this.parseInserted(el, (e) => this.parseParagraph(e)));
          break;

        case "del":
          result.children.push(this.parseDeleted(el, (e) => this.parseParagraph(e)));
          break;
      }
    }

    return result;
  }

  parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
    this.parseDefaultProperties(elem, (paragraph.cssStyle = {}), null, (c) => {
      if (parseParagraphProperty(c, paragraph, this.xmlParser)) return true;

      switch (c.localName) {
        case "pStyle":
          paragraph.styleName = this.xmlParser.attr(c, "val");
          break;

        case "cnfStyle":
          paragraph.className = values.classNameOfCnfStyle(c);
          break;

        case "framePr":
          this.parseFrame(c, paragraph);
          break;

        case "rPr":
          //TODO ignore
          break;

        default:
          return false;
      }

      return true;
    });
  }

  parseFrame(node: Element, paragraph: WmlParagraph) {
    const dropCap = this.xmlParser.attr(node, "dropCap");

    if (dropCap == "drop") paragraph.cssStyle["float"] = "left";
  }

  parseHyperlink(node: Element, parent?: OpenXmlElement): WmlHyperlink {
    const result: WmlHyperlink = {
      type: "hyperlink",
      parent: parent,
      children: [],
    } as WmlHyperlink;

    result.anchor = this.xmlParser.attr(node, "anchor");
    result.id = this.xmlParser.attr(node, "id");

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "r":
          result.children.push(this.parseRun(c, result));
          break;
      }
    });

    return result;
  }

  parseSmartTag(node: Element, parent?: OpenXmlElement): WmlSmartTag {
    const result: WmlSmartTag = { type: "smartTag", parent, children: [] };
    const uri = this.xmlParser.attr(node, "uri");
    const element = this.xmlParser.attr(node, "element");

    if (uri) result.uri = uri;

    if (element) result.element = element;

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "r":
          result.children.push(this.parseRun(c, result));
          break;
      }
    });

    return result;
  }

  parseRun(node: Element, parent?: OpenXmlElement): WmlRun {
    const result: WmlRun = { type: "run", parent: parent, children: [] } as WmlRun;

    xmlUtil.foreach(node, (c) => {
      c = this.checkAlternateContent(c);

      switch (c.localName) {
        case "t":
          result.children.push({
            type: "text",
            text: c.textContent,
          } as WmlText); //.replace(" ", "\u00A0"); // TODO
          break;

        case "delText":
          result.children.push({
            type: "deletedText",
            text: c.textContent,
          } as WmlText);
          break;

        case "commentReference":
          result.children.push(new WmlCommentReference(this.xmlParser.attr(c, "id")));
          break;

        case "fldSimple":
          result.children.push({
            type: "simpleField",
            instruction: this.xmlParser.attr(c, "instr"),
            lock: this.xmlParser.boolAttr(c, "lock", false),
            dirty: this.xmlParser.boolAttr(c, "dirty", false),
          } as WmlFieldSimple);
          break;

        case "instrText":
          result.fieldRun = true;
          result.children.push({
            type: "instruction",
            text: c.textContent,
          } as WmlInstructionText);
          break;

        case "fldChar":
          result.fieldRun = true;
          result.children.push({
            type: "complexField",
            charType: this.xmlParser.attr(c, "fldCharType"),
            lock: this.xmlParser.boolAttr(c, "lock", false),
            dirty: this.xmlParser.boolAttr(c, "dirty", false),
          } as WmlFieldChar);
          break;

        case "noBreakHyphen":
          result.children.push({ type: "noBreakHyphen" });
          break;

        case "br":
          result.children.push({
            type: "break",
            break: this.xmlParser.attr(c, "type") || "textWrapping",
          } as WmlBreak);
          break;

        case "lastRenderedPageBreak":
          result.children.push({
            type: "break",
            break: "lastRenderedPageBreak",
          } as WmlBreak);
          break;

        case "sym":
          result.children.push({
            type: "symbol",
            font: encloseFontFamily(this.xmlParser.attr(c, "font")),
            char: this.xmlParser.attr(c, "char"),
          } as WmlSymbol);
          break;

        case "tab":
          result.children.push({ type: "tab" });
          break;

        case "footnoteReference":
          result.children.push({
            type: "footnoteReference",
            id: this.xmlParser.attr(c, "id"),
          } as WmlNoteReference);
          break;

        case "endnoteReference":
          result.children.push({
            type: "endnoteReference",
            id: this.xmlParser.attr(c, "id"),
          } as WmlNoteReference);
          break;

        case "drawing":
          const d = this.parseDrawing(c);

          if (d) result.children = [d];
          break;

        case "pict":
          result.children.push(this.parseVmlPicture(c));
          break;

        case "rPr":
          this.parseRunProperties(c, result);
          break;
      }
    });

    return result;
  }

  parseMathElement(elem: Element): OpenXmlElement {
    const propsTag = `${elem.localName}Pr`;
    const result = { type: mmlTagMap[elem.localName], children: [] } as OpenXmlElement;

    for (const el of this.xmlParser.elements(elem)) {
      const childType = mmlTagMap[el.localName];

      if (childType) {
        result.children.push(this.parseMathElement(el));
      } else if (el.localName == "r") {
        const run = this.parseRun(el);
        run.type = "mmlRun";
        result.children.push(run);
      } else if (el.localName == propsTag) {
        result.props = this.parseMathProperies(el);
      }
    }

    return result;
  }

  parseMathProperies(elem: Element): Record<string, string | boolean | number> {
    const result: Record<string, string | boolean | number> = {};

    for (const el of this.xmlParser.elements(elem)) {
      switch (el.localName) {
        case "chr":
          result.char = this.xmlParser.attr(el, "val");
          break;
        case "vertJc":
          result.verticalJustification = this.xmlParser.attr(el, "val");
          break;
        case "pos":
          result.position = this.xmlParser.attr(el, "val");
          break;
        case "degHide":
          result.hideDegree = this.xmlParser.boolAttr(el, "val");
          break;
        case "begChr":
          result.beginChar = this.xmlParser.attr(el, "val");
          break;
        case "endChr":
          result.endChar = this.xmlParser.attr(el, "val");
          break;
      }
    }

    return result;
  }

  parseRunProperties(elem: Element, run: WmlRun) {
    this.parseDefaultProperties(elem, (run.cssStyle = {}), null, (c) => {
      switch (c.localName) {
        case "rStyle":
          run.styleName = this.xmlParser.attr(c, "val");
          break;

        case "vertAlign":
          run.verticalAlign = values.valueOfVertAlign(c, true);
          break;

        default:
          return false;
      }

      return true;
    });
  }

  parseVmlPicture(elem: Element): OpenXmlElement {
    const result = { type: "vmlPicture" as DomType, children: [] };

    for (const el of this.xmlParser.elements(elem)) {
      const child = parseVmlElement(el, this);
      if (child) {
        result.children.push(child);
      }
    }

    return result;
  }

  checkAlternateContent(elem: Element): Element {
    if (elem.localName != "AlternateContent") return elem;

    const choice = this.xmlParser.element(elem, "Choice");

    if (choice) {
      const requires = this.xmlParser.attr(choice, "Requires");
      const namespaceURI = elem.lookupNamespaceURI(requires);

      if (supportedNamespaceURIs.includes(namespaceURI)) return choice.firstElementChild;
    }

    return this.xmlParser.element(elem, "Fallback")?.firstElementChild;
  }

  parseDrawing(node: Element): OpenXmlElement {
    for (const n of this.xmlParser.elements(node)) {
      switch (n.localName) {
        case "inline":
        case "anchor":
          return this.parseDrawingWrapper(n);
      }
    }
  }

  parseDrawingWrapper(node: Element): OpenXmlElement {
    const result = { type: "drawing", children: [], cssStyle: {} } as OpenXmlElement;
    const isAnchor = node.localName == "anchor";

    //TODO
    // result.style["margin-left"] = xml.sizeAttr(node, "distL", SizeType.Emu);
    // result.style["margin-top"] = xml.sizeAttr(node, "distT", SizeType.Emu);
    // result.style["margin-right"] = xml.sizeAttr(node, "distR", SizeType.Emu);
    // result.style["margin-bottom"] = xml.sizeAttr(node, "distB", SizeType.Emu);

    let wrapType: Nullable<"wrapTopAndBottom" | "wrapNone"> = null;
    const simplePos = this.xmlParser.boolAttr(node, "simplePos");

    const posX = { relative: "page", align: "left", offset: "0" };
    const posY = { relative: "page", align: "top", offset: "0" };

    for (const n of this.xmlParser.elements(node)) {
      switch (n.localName) {
        case "simplePos":
          if (simplePos) {
            posX.offset = this.xmlParser.lengthAttr(n, "x", LengthUsage.Emu);
            posY.offset = this.xmlParser.lengthAttr(n, "y", LengthUsage.Emu);
          }
          break;

        case "extent":
          result.cssStyle["width"] = this.xmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
          result.cssStyle["height"] = this.xmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
          break;

        case "positionH":
        case "positionV":
          if (!simplePos) {
            const pos = n.localName == "positionH" ? posX : posY;
            const alignNode = this.xmlParser.element(n, "align");
            const offsetNode = this.xmlParser.element(n, "posOffset");

            pos.relative = this.xmlParser.attr(n, "relativeFrom") ?? pos.relative;

            if (alignNode) pos.align = alignNode.textContent;

            if (offsetNode) pos.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
          }
          break;

        case "wrapTopAndBottom":
          wrapType = "wrapTopAndBottom";
          break;

        case "wrapNone":
          wrapType = "wrapNone";
          break;

        case "graphic":
          var g = this.parseGraphic(n);

          if (g) result.children.push(g);
          break;
      }
    }

    if (wrapType == "wrapTopAndBottom") {
      result.cssStyle["display"] = "block";

      if (posX.align) {
        result.cssStyle["text-align"] = posX.align;
        result.cssStyle["width"] = "100%";
      }
    } else if (wrapType == "wrapNone") {
      result.cssStyle["display"] = "block";
      result.cssStyle["position"] = "relative";
      result.cssStyle["width"] = "0px";
      result.cssStyle["height"] = "0px";

      if (posX.offset) result.cssStyle["left"] = posX.offset;
      if (posY.offset) result.cssStyle["top"] = posY.offset;
    } else if (isAnchor && (posX.align == "left" || posX.align == "right")) {
      result.cssStyle["float"] = posX.align;
    }

    return result;
  }

  parseGraphic(elem: Element): OpenXmlElement {
    const graphicData = this.xmlParser.element(elem, "graphicData");

    for (const n of this.xmlParser.elements(graphicData)) {
      switch (n.localName) {
        case "pic":
          return this.parsePicture(n);
      }
    }

    return null;
  }

  parsePicture(elem: Element): IDomImage {
    const result = { type: "image", src: "", cssStyle: {} } as IDomImage;
    const blipFill = this.xmlParser.element(elem, "blipFill");
    const blip = this.xmlParser.element(blipFill, "blip");

    result.src = this.xmlParser.attr(blip, "embed");

    const spPr = this.xmlParser.element(elem, "spPr");
    const xfrm = this.xmlParser.element(spPr, "xfrm");

    result.cssStyle["position"] = "relative";

    for (const n of this.xmlParser.elements(xfrm)) {
      switch (n.localName) {
        case "ext":
          result.cssStyle["width"] = this.xmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
          result.cssStyle["height"] = this.xmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
          break;

        case "off":
          result.cssStyle["left"] = this.xmlParser.lengthAttr(n, "x", LengthUsage.Emu);
          result.cssStyle["top"] = this.xmlParser.lengthAttr(n, "y", LengthUsage.Emu);
          break;
      }
    }

    return result;
  }

  parseTable(node: Element): WmlTable {
    const result: WmlTable = { type: "table", children: [] };

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "tr":
          result.children.push(this.parseTableRow(c));
          break;

        case "tblGrid":
          result.columns = this.parseTableColumns(c);
          break;

        case "tblPr":
          this.parseTableProperties(c, result);
          break;
      }
    });

    return result;
  }

  parseTableColumns(node: Element): WmlTableColumn[] {
    const result = [];

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case "gridCol":
          result.push({ width: this.xmlParser.lengthAttr(n, "w") });
          break;
      }
    });

    return result;
  }

  parseTableProperties(elem: Element, table: WmlTable) {
    table.cssStyle = {};
    table.cellStyle = {};

    this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, (c) => {
      switch (c.localName) {
        case "tblStyle":
          table.styleName = this.xmlParser.attr(c, "val");
          break;

        case "tblLook":
          table.className = values.classNameOftblLook(c);
          break;

        case "tblpPr":
          this.parseTablePosition(c, table);
          break;

        case "tblStyleColBandSize":
          table.colBandSize = this.xmlParser.intAttr(c, "val");
          break;

        case "tblStyleRowBandSize":
          table.rowBandSize = this.xmlParser.intAttr(c, "val");
          break;

        case "hidden":
          table.cssStyle["display"] = "none";
          break;

        default:
          return false;
      }

      return true;
    });

    switch (table.cssStyle["text-align"]) {
      case "center":
        delete table.cssStyle["text-align"];
        table.cssStyle["margin-left"] = "auto";
        table.cssStyle["margin-right"] = "auto";
        break;

      case "right":
        delete table.cssStyle["text-align"];
        table.cssStyle["margin-left"] = "auto";
        break;
    }
  }

  parseTablePosition(node: Element, table: WmlTable) {
    const topFromText = this.xmlParser.lengthAttr(node, "topFromText");
    const bottomFromText = this.xmlParser.lengthAttr(node, "bottomFromText");
    const rightFromText = this.xmlParser.lengthAttr(node, "rightFromText");
    const leftFromText = this.xmlParser.lengthAttr(node, "leftFromText");

    table.cssStyle["float"] = "left";
    table.cssStyle["margin-bottom"] = values.addSize(
      table.cssStyle["margin-bottom"],
      bottomFromText,
    );
    table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
    table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
    table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
  }

  parseTableRow(node: Element): WmlTableRow {
    const result: WmlTableRow = { type: "row", children: [] };

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "tc":
          result.children.push(this.parseTableCell(c));
          break;

        case "trPr":
          this.parseTableRowProperties(c, result);
          break;
      }
    });

    return result;
  }

  parseTableRowProperties(elem: Element, row: WmlTableRow) {
    row.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
      switch (c.localName) {
        case "cnfStyle":
          row.className = values.classNameOfCnfStyle(c);
          break;

        case "tblHeader":
          row.isHeader = this.xmlParser.boolAttr(c, "val");
          break;

        case "gridBefore":
          row.gridBefore = this.xmlParser.intAttr(c, "val");
          break;

        case "gridAfter":
          row.gridAfter = this.xmlParser.intAttr(c, "val");
          break;

        default:
          return false;
      }

      return true;
    });
  }

  parseTableCell(node: Element): OpenXmlElement {
    const result: WmlTableCell = { type: "cell", children: [] };

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "tbl":
          result.children.push(this.parseTable(c));
          break;

        case "p":
          result.children.push(this.parseParagraph(c));
          break;

        case "tcPr":
          this.parseTableCellProperties(c, result);
          break;
      }
    });

    return result;
  }

  parseTableCellProperties(elem: Element, cell: WmlTableCell) {
    cell.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
      switch (c.localName) {
        case "gridSpan":
          cell.span = this.xmlParser.intAttr(c, "val", null);
          break;

        case "vMerge":
          cell.verticalMerge = this.xmlParser.attr(c, "val") ?? "continue";
          break;

        case "cnfStyle":
          cell.className = values.classNameOfCnfStyle(c);
          break;

        default:
          return false;
      }

      return true;
    });

    this.parseTableCellVerticalText(elem, cell);
  }

  parseTableCellVerticalText(elem: Element, cell: WmlTableCell) {
    const directionMap = {
      btLr: {
        writingMode: "vertical-rl",
        transform: "rotate(180deg)",
      },
      lrTb: {
        writingMode: "vertical-lr",
        transform: "none",
      },
      tbRl: {
        writingMode: "vertical-rl",
        transform: "none",
      },
    };

    xmlUtil.foreach(elem, (c) => {
      if (c.localName === "textDirection") {
        const direction = this.xmlParser.attr(c, "val");
        const style = directionMap[direction] || { writingMode: "horizontal-tb" };
        cell.cssStyle["writing-mode"] = style.writingMode;
        cell.cssStyle["transform"] = style.transform;
      }
    });
  }

  parseDefaultProperties(
    elem: Element,
    style: Nullable<Record<string, string>> = null,
    childStyle: Nullable<Record<string, string>> = null,
    handler: Nullable<(prop: Element) => boolean> = null,
  ): Record<string, string> {
    style = style || {};

    xmlUtil.foreach(elem, (c) => {
      if (handler?.(c)) return;

      switch (c.localName) {
        case "jc":
          style["text-align"] = values.valueOfJc(c);
          break;

        case "textAlignment":
          style["vertical-align"] = values.valueOfTextAlignment(c);
          break;

        case "color":
          style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
          break;

        case "sz":
          style["font-size"] = style["min-height"] = this.xmlParser.lengthAttr(
            c,
            "val",
            LengthUsage.FontSize,
          );
          break;

        case "shd":
          style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
          break;

        case "highlight":
          style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
          break;

        case "vertAlign":
          //TODO
          // style.verticalAlign = values.valueOfVertAlign(c);
          break;

        case "position":
          style.verticalAlign = this.xmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
          break;

        case "tcW":
          if (this.options.ignoreWidth) break;

        case "tblW":
          style["width"] = values.valueOfSize(c, "w");
          break;

        case "trHeight":
          this.parseTrHeight(c, style);
          break;

        case "strike":
          style["text-decoration"] = this.xmlParser.boolAttr(c, "val", true)
            ? "line-through"
            : "none";
          break;

        case "b":
          style["font-weight"] = this.xmlParser.boolAttr(c, "val", true) ? "bold" : "normal";
          break;

        case "i":
          style["font-style"] = this.xmlParser.boolAttr(c, "val", true) ? "italic" : "normal";
          break;

        case "caps":
          style["text-transform"] = this.xmlParser.boolAttr(c, "val", true) ? "uppercase" : "none";
          break;

        case "smallCaps":
          style["font-variant"] = this.xmlParser.boolAttr(c, "val", true) ? "small-caps" : "none";
          break;

        case "u":
          this.parseUnderline(c, style);
          break;

        case "ind":
        case "tblInd":
          this.parseIndentation(c, style);
          break;

        case "rFonts":
          this.parseFont(c, style);
          break;

        case "tblBorders":
          this.parseBorderProperties(c, childStyle || style);
          break;

        case "tblCellSpacing":
          style["border-spacing"] = values.valueOfMargin(c);
          style["border-collapse"] = "separate";
          break;

        case "pBdr":
          this.parseBorderProperties(c, style);
          break;

        case "bdr":
          style["border"] = values.valueOfBorder(c);
          break;

        case "tcBorders":
          this.parseBorderProperties(c, style);
          break;

        case "vanish":
          if (this.xmlParser.boolAttr(c, "val", true)) style["display"] = "none";
          break;

        case "kern":
          //TODO
          //style['letter-spacing'] = xml.lengthAttr(elem, 'val', LengthUsage.FontSize);
          break;

        case "noWrap":
          //TODO
          //style["white-space"] = "nowrap";
          break;

        case "tblCellMar":
        case "tcMar":
          this.parseMarginProperties(c, childStyle || style);
          break;

        case "tblLayout":
          style["table-layout"] = values.valueOfTblLayout(c);
          break;

        case "vAlign":
          style["vertical-align"] = values.valueOfTextAlignment(c);
          break;

        case "spacing":
          if (elem.localName == "pPr") this.parseSpacing(c, style);
          break;

        case "wordWrap":
          if (this.xmlParser.boolAttr(c, "val"))
            //TODO: test with examples
            style["overflow-wrap"] = "break-word";
          break;

        case "suppressAutoHyphens":
          style["hyphens"] = this.xmlParser.boolAttr(c, "val", true) ? "none" : "auto";
          break;

        case "lang":
          style["$lang"] = this.xmlParser.attr(c, "val");
          break;

        case "bCs":
        case "iCs":
        case "szCs":
        case "tabs": //ignore - tabs is parsed by other parser
        case "outlineLvl": //TODO
        case "contextualSpacing": //TODO
        case "tblStyleColBandSize": //TODO
        case "tblStyleRowBandSize": //TODO
        case "webHidden": //TODO - maybe web-hidden should be implemented
        case "pageBreakBefore": //TODO - maybe ignore
        case "suppressLineNumbers": //TODO - maybe ignore
        case "keepLines": //TODO - maybe ignore
        case "keepNext": //TODO - maybe ignore
        case "widowControl": //TODO - maybe ignore
        case "bidi": //TODO - maybe ignore
        case "rtl": //TODO - maybe ignore
        case "noProof": //ignore spellcheck
          //TODO ignore
          break;

        default:
          if (this.options.debug) {
            // eslint-disable-next-line no-console
            console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
          }
          break;
      }
    });

    return style;
  }

  parseUnderline(node: Element, style: Record<string, string>) {
    const val = this.xmlParser.attr(node, "val");

    if (val == null) return;

    switch (val) {
      case "dash":
      case "dashDotDotHeavy":
      case "dashDotHeavy":
      case "dashedHeavy":
      case "dashLong":
      case "dashLongHeavy":
      case "dotDash":
      case "dotDotDash":
        style["text-decoration"] = "underline dashed";
        break;

      case "dotted":
      case "dottedHeavy":
        style["text-decoration"] = "underline dotted";
        break;

      case "double":
        style["text-decoration"] = "underline double";
        break;

      case "single":
      case "thick":
        style["text-decoration"] = "underline";
        break;

      case "wave":
      case "wavyDouble":
      case "wavyHeavy":
        style["text-decoration"] = "underline wavy";
        break;

      case "words":
        style["text-decoration"] = "underline";
        break;

      case "none":
        style["text-decoration"] = "none";
        break;
    }

    const col = xmlUtil.colorAttr(node, "color");

    if (col) style["text-decoration-color"] = col;
  }

  parseFont(node: Element, style: Record<string, string>) {
    const ascii = this.xmlParser.attr(node, "ascii");
    const asciiTheme = values.themeValue(node, "asciiTheme");
    const eastAsia = this.xmlParser.attr(node, "eastAsia");
    const fonts = [ascii, asciiTheme, eastAsia].filter((x) => x).map((x) => encloseFontFamily(x));

    if (fonts.length > 0) style["font-family"] = [...new Set(fonts)].join(", ");
  }

  parseIndentation(node: Element, style: Record<string, string>) {
    const firstLine = this.xmlParser.lengthAttr(node, "firstLine");
    const hanging = this.xmlParser.lengthAttr(node, "hanging");
    const left = this.xmlParser.lengthAttr(node, "left");
    const start = this.xmlParser.lengthAttr(node, "start");
    const right = this.xmlParser.lengthAttr(node, "right");
    const end = this.xmlParser.lengthAttr(node, "end");

    if (firstLine) style["text-indent"] = firstLine;
    if (hanging) style["text-indent"] = `-${hanging}`;
    if (left || start) style["margin-left"] = left || start;
    if (right || end) style["margin-right"] = right || end;
  }

  parseSpacing(node: Element, style: Record<string, string>) {
    const before = this.xmlParser.lengthAttr(node, "before");
    const after = this.xmlParser.lengthAttr(node, "after");
    const line = this.xmlParser.intAttr(node, "line", null);
    const lineRule = this.xmlParser.attr(node, "lineRule");

    if (before) style["margin-top"] = before;
    if (after) style["margin-bottom"] = after;

    if (line !== null) {
      switch (lineRule) {
        case "auto":
          style["line-height"] = `${(line / 240).toFixed(2)}`;
          break;

        case "atLeast":
          style["line-height"] = `calc(100% + ${line / 20}pt)`;
          break;

        default:
          style["line-height"] = style["min-height"] = `${line / 20}pt`;
          break;
      }
    }
  }

  parseMarginProperties(node: Element, output: Record<string, string>) {
    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "left":
          output["padding-left"] = values.valueOfMargin(c);
          break;

        case "right":
          output["padding-right"] = values.valueOfMargin(c);
          break;

        case "top":
          output["padding-top"] = values.valueOfMargin(c);
          break;

        case "bottom":
          output["padding-bottom"] = values.valueOfMargin(c);
          break;
      }
    });
  }

  parseTrHeight(node: Element, output: Record<string, string>) {
    switch (this.xmlParser.attr(node, "hRule")) {
      case "exact":
        output["height"] = this.xmlParser.lengthAttr(node, "val");
        break;

      case "atLeast":
      default:
        output["height"] = this.xmlParser.lengthAttr(node, "val");
        // min-height doesn't work for tr
        //output["min-height"] = this.xmlParser.sizeAttr(node, "val");
        break;
    }
  }

  parseBorderProperties(node: Element, output: Record<string, string>) {
    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case "start":
        case "left":
          output["border-left"] = values.valueOfBorder(c);
          break;

        case "end":
        case "right":
          output["border-right"] = values.valueOfBorder(c);
          break;

        case "top":
          output["border-top"] = values.valueOfBorder(c);
          break;

        case "bottom":
          output["border-bottom"] = values.valueOfBorder(c);
          break;
      }
    });
  }
}

const knownColors = [
  "black",
  "blue",
  "cyan",
  "darkBlue",
  "darkCyan",
  "darkGray",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "green",
  "lightGray",
  "magenta",
  "none",
  "red",
  "white",
  "yellow",
];

class xmlUtil {
  static foreach(node: Element, cb: (n: Element) => void) {
    for (let i = 0; i < node.childNodes.length; i++) {
      const n = node.childNodes[i];

      if (n.nodeType == Node.ELEMENT_NODE) cb(n as Element);
    }
  }

  static colorAttr(
    node: Element,
    attrName: string,
    defValue: Nullable<string> = null,
    autoColor = "black",
  ) {
    const xmlParser = new XmlParser();
    const v = xmlParser.attr(node, attrName);

    if (v) {
      if (v == "auto") {
        return autoColor;
      } else if (knownColors.includes(v)) {
        return v;
      }

      return `#${v}`;
    }

    const themeColor = xmlParser.attr(node, "themeColor");

    return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
  }

  static sizeValue(node: Element, type: LengthUsageType = LengthUsage.Dxa) {
    return convertLength(node.textContent, type);
  }
}

class values {
  static themeValue(c: Element, attr: string) {
    const xmlParser = new XmlParser();
    const val = xmlParser.attr(c, attr);
    return val ? `var(--docx-${val}-font)` : null;
  }

  static valueOfSize(c: Element, attr: string) {
    const xmlParser = new XmlParser();
    let type = LengthUsage.Dxa;

    switch (xmlParser.attr(c, "type")) {
      case "dxa":
        break;
      case "pct":
        type = LengthUsage.Percent;
        break;
      case "auto":
        return "auto";
    }

    return xmlParser.lengthAttr(c, attr, type);
  }

  static valueOfMargin(c: Element) {
    const xmlParser = new XmlParser();
    return xmlParser.lengthAttr(c, "w");
  }

  static valueOfBorder(c: Element) {
    const xmlParser = new XmlParser();
    const type = values.parseBorderType(xmlParser.attr(c, "val"));

    if (type == "none") return "none";

    const color = xmlUtil.colorAttr(c, "color");
    const size = xmlParser.lengthAttr(c, "sz", LengthUsage.Border);

    return `${size} ${type} ${color == "auto" ? autos.borderColor : color}`;
  }

  static parseBorderType(type: string) {
    switch (type) {
      case "single":
        return "solid";
      case "dashDotStroked":
        return "solid";
      case "dashed":
        return "dashed";
      case "dashSmallGap":
        return "dashed";
      case "dotDash":
        return "dotted";
      case "dotDotDash":
        return "dotted";
      case "dotted":
        return "dotted";
      case "double":
        return "double";
      case "doubleWave":
        return "double";
      case "inset":
        return "inset";
      case "nil":
        return "none";
      case "none":
        return "none";
      case "outset":
        return "outset";
      case "thick":
        return "solid";
      case "thickThinLargeGap":
        return "solid";
      case "thickThinMediumGap":
        return "solid";
      case "thickThinSmallGap":
        return "solid";
      case "thinThickLargeGap":
        return "solid";
      case "thinThickMediumGap":
        return "solid";
      case "thinThickSmallGap":
        return "solid";
      case "thinThickThinLargeGap":
        return "solid";
      case "thinThickThinMediumGap":
        return "solid";
      case "thinThickThinSmallGap":
        return "solid";
      case "threeDEmboss":
        return "solid";
      case "threeDEngrave":
        return "solid";
      case "triple":
        return "double";
      case "wave":
        return "solid";
    }

    return "solid";
  }

  static valueOfTblLayout(c: Element) {
    const xmlParser = new XmlParser();
    const type = xmlParser.attr(c, "val");
    return type == "fixed" ? "fixed" : "auto";
  }

  static classNameOfCnfStyle(c: Element) {
    const xmlParser = new XmlParser();
    const val = xmlParser.attr(c, "val");
    const classes = [
      "first-row",
      "last-row",
      "first-col",
      "last-col",
      "odd-col",
      "even-col",
      "odd-row",
      "even-row",
      "ne-cell",
      "nw-cell",
      "se-cell",
      "sw-cell",
    ];

    return classes.filter((_, i) => val[i] == "1").join(" ");
  }

  static valueOfJc(c: Element) {
    const xmlParser = new XmlParser();
    const type = xmlParser.attr(c, "val");

    switch (type) {
      case "start":
      case "left":
        return "left";
      case "center":
        return "center";
      case "end":
      case "right":
        return "right";
      case "both":
        return "justify";
    }

    return type;
  }

  static valueOfVertAlign(c: Element, asTagName = false) {
    const xmlParser = new XmlParser();
    const type = xmlParser.attr(c, "val");

    switch (type) {
      case "subscript":
        return "sub";
      case "superscript":
        return asTagName ? "sup" : "super";
    }

    return asTagName ? null : type;
  }

  static valueOfTextAlignment(c: Element) {
    const xmlParser = new XmlParser();
    const type = xmlParser.attr(c, "val");

    switch (type) {
      case "auto":
      case "baseline":
        return "baseline";
      case "top":
        return "top";
      case "center":
        return "middle";
      case "bottom":
        return "bottom";
    }

    return type;
  }

  static addSize(a: string, b: string): string {
    if (a == null) return b;
    if (b == null) return a;

    return `calc(${a} + ${b})`; //TODO
  }

  static classNameOftblLook(c: Element) {
    const xmlParser = new XmlParser();
    const val = xmlParser.hexAttr(c, "val", 0);
    let className = "";

    if (xmlParser.boolAttr(c, "firstRow") || val & 0x0020) className += " first-row";
    if (xmlParser.boolAttr(c, "lastRow") || val & 0x0040) className += " last-row";
    if (xmlParser.boolAttr(c, "firstColumn") || val & 0x0080) className += " first-col";
    if (xmlParser.boolAttr(c, "lastColumn") || val & 0x0100) className += " last-col";
    if (xmlParser.boolAttr(c, "noHBand") || val & 0x0200) className += " no-hband";
    if (xmlParser.boolAttr(c, "noVBand") || val & 0x0400) className += " no-vband";

    return className.trim();
  }
}

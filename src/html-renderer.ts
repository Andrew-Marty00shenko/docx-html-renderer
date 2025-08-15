import type { WmlComment, WmlCommentRangeStart, WmlCommentReference } from "./comments/elements";
import type { Part } from "./common/part";
import type { WmlBookmarkStart } from "./document/bookmarks";
import type { CommonProperties } from "./document/common";
import { calculateTotalElementHeight, getComputedStyles, pxToPt } from "./document/common";
import type { DocumentElement } from "./document/document";
import type {
  IDomImage,
  IDomNumbering,
  OpenXmlElement,
  WmlAltChunk,
  WmlBreak,
  WmlHyperlink,
  WmlNoteReference,
  WmlSmartTag,
  WmlSymbol,
  WmlTable,
  WmlTableCell,
  WmlTableColumn,
  WmlTableRow,
  WmlText,
} from "./document/dom";
import { DomType } from "./document/dom";
import type { WmlParagraph } from "./document/paragraph";
import type { RunProperties, WmlRun } from "./document/run";
import type { FooterHeaderReference, SectionProperties } from "./document/section";
import type { IDomStyle } from "./document/style";
import type { FontTablePart } from "./font-table/font-table";
import type { BaseHeaderFooterPart } from "./header-footer/parts";
import type { WmlBaseNote, WmlFootnote } from "./notes/elements";
import type { ThemePart } from "./theme/theme-part";
import type { VmlElement } from "./vml/vml";
import type { Options } from "./docx2html";
import { computePixelToPoint, updateTabStop } from "./javascript";
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep } from "./utils";
import type { WordDocument } from "./word-document";

const ns = {
  svg: "http://www.w3.org/2000/svg",
  mathML: "http://www.w3.org/1998/Math/MathML",
};

interface CellPos {
  col: number;
  row: number;
}

interface Section {
  sectProps: SectionProperties;
  elements: OpenXmlElement[];
  pageBreak: boolean;
}

declare const Highlight: any;

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

interface GetBreaksSectionsProps {
  article: Element;
  availableHeight: number;
  sections: HTMLElement[];
  section: HTMLElement;
  sectionIndex: number;
  bodyContainer: HTMLElement;
  sectionsProps: SectionProperties[];
}

interface GetTableProps {
  currentTable: HTMLElement;
  articleChildrenHeight: number;
  availableHeight: number;
}

export class HtmlRenderer {
  // TODO remove hardcode with null!
  className = "docx";
  rootSelector = "";
  document: WordDocument = null!;
  options: Options = {} as Options;
  styleMap: Record<string, IDomStyle> = {};
  currentPart: Part | null = null;

  tableVerticalMerges: CellVerticalMergeType[] = [];
  currentVerticalMerge: CellVerticalMergeType = null!;
  tableCellPositions: CellPos[] = [];
  currentCellPosition: CellPos = null!;

  footnoteMap: Record<string, WmlFootnote> = {};
  endnoteMap: Record<string, WmlFootnote> = {};
  currentFootnoteIds: string[] = [];
  currentEndnoteIds: string[] = [];
  usedHeaderFooterParts: any[] = [];

  defaultTabSize = "";
  currentTabs: any[] = [];

  commentHighlight: any;
  commentMap: Record<string, Range> = {};

  tasks: Promise<any>[] = [];
  postRenderTasks: any[] = [];

  constructor(public htmlDocument: Document) {}

  async render(
    document: WordDocument,
    bodyContainer: HTMLElement,
    styleContainer: HTMLElement | null = null,
    options: Options,
  ) {
    this.document = document;
    this.options = options;
    this.className = options.className;
    this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ":root";
    this.styleMap = null!;
    this.tasks = [];

    if (this.options.renderComments && globalThis.Highlight) {
      this.commentHighlight = new Highlight();
    }

    styleContainer = styleContainer || bodyContainer;

    removeAllElements(styleContainer);
    removeAllElements(bodyContainer);

    styleContainer.appendChild(this.createComment("docxjs library predefined styles"));
    styleContainer.appendChild(this.renderDefaultStyle());

    if (document.themePart) {
      styleContainer.appendChild(this.createComment("docxjs document theme values"));
      this.renderTheme(document.themePart, styleContainer);
    }

    if (document.stylesPart != null) {
      this.styleMap = this.processStyles(document.stylesPart.styles);

      styleContainer.appendChild(this.createComment("docxjs document styles"));
      styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
    }

    if (document.numberingPart) {
      this.prodessNumberings(document.numberingPart.domNumberings);

      styleContainer.appendChild(this.createComment("docxjs document numbering styles"));
      styleContainer.appendChild(
        this.renderNumbering(document.numberingPart.domNumberings, styleContainer),
      );
    }

    if (document.footnotesPart) {
      this.footnoteMap = keyBy(document.footnotesPart.notes, (item) => item.id);
    }

    if (document.endnotesPart) {
      this.endnoteMap = keyBy(document.endnotesPart.notes, (item) => item.id);
    }

    if (document.settingsPart) {
      this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
    }

    if (!options.ignoreFonts && document.fontTablePart)
      this.renderFontTable(document.fontTablePart, styleContainer);

    const { result: sectionElements, sectionsProps } = this.renderSections(
      document.documentPart.body,
    );

    if (this.options.inWrapper) {
      bodyContainer.appendChild(this.renderWrapper(sectionElements));

      /** call breaking child sections */
      if (this.options.breakPages && this.options.ignoreLastRenderedPageBreak) {
        this.renderChildSections(sectionElements, bodyContainer, sectionsProps);
      }
    } else {
      appendChildren(bodyContainer, sectionElements);
    }

    if (this.commentHighlight && options.renderComments) {
      (CSS as any).highlights.set(`${this.className}-comments`, this.commentHighlight);
    }

    this.postRenderTasks.forEach((task) => task());

    await Promise.allSettled(this.tasks);

    this.refreshTabStops();
  }

  getModifiedTables({ articleChildrenHeight, currentTable, availableHeight }: GetTableProps) {
    let breakingChildIndex = 0;
    let newTable: HTMLElement | null = null;
    let hasOversizeElement = false;

    for (let index = 0; index < currentTable.children.length; index++) {
      const tableChildNode = currentTable.children[index] as Element;

      if (tableChildNode.localName === "colgroup") {
        continue;
      }

      const totalElementHeight = calculateTotalElementHeight(tableChildNode as Element);

      articleChildrenHeight += totalElementHeight;

      if (articleChildrenHeight >= availableHeight) {
        if (totalElementHeight > availableHeight - articleChildrenHeight) {
          hasOversizeElement = true;
          breakingChildIndex = index + 1;

          break;
        }

        breakingChildIndex = index;

        break;
      }
    }

    const nodesToKeep = Array.from(currentTable.children).slice(
      0,
      breakingChildIndex === 0 ? currentTable.children.length : breakingChildIndex,
    );
    const nodesToMove =
      breakingChildIndex === 0 ? [] : Array.from(currentTable.children).slice(breakingChildIndex);

    currentTable.replaceChildren(...nodesToKeep);

    if (nodesToMove.length > 0) {
      newTable = this.createElement("table");

      Array.from(currentTable.attributes).forEach((attr) => {
        newTable?.setAttribute(attr.name, attr.value);
      });

      const colgroup = currentTable.querySelector("colgroup");

      if (colgroup) {
        newTable.appendChild(colgroup.cloneNode(true));
      }

      nodesToMove.forEach((node) => newTable?.appendChild(node));
    }

    return { articleChildrenHeight, newTable, currentTable, hasOversizeElement };
  }

  getBreaksSections({
    article,
    availableHeight,
    bodyContainer,
    section,
    sectionIndex,
    sections,
    sectionsProps,
  }: GetBreaksSectionsProps) {
    /** Index of the section element after which the content (that didn't fit in the current section) will be inserted. */
    let replaceSectionIndex = 0;

    /** Index from which the content of the Article will be truncated. */
    let breakingChildIndex = 0;

    let articleChildrenHeight = 0;

    let newSection: HTMLElement | null = null;
    let currentTable: HTMLElement | null = null;
    let newTable: HTMLElement | null = null;

    let hasOversizeElement = false;

    for (let index = 0; index < article.children.length; index++) {
      const articleChildNode = article.children[index] as HTMLElement;

      if (articleChildNode.localName === "table") {
        const modifiedTables = this.getModifiedTables({
          currentTable: articleChildNode,
          articleChildrenHeight,
          availableHeight,
        });

        articleChildrenHeight += modifiedTables.articleChildrenHeight;
        currentTable = modifiedTables.currentTable;
        newTable = modifiedTables.newTable;
        hasOversizeElement = modifiedTables.hasOversizeElement;

        if (articleChildrenHeight >= availableHeight) {
          breakingChildIndex = index + 1;
          break;
        }
      } else {
        const totalElementHeight = calculateTotalElementHeight(articleChildNode);

        articleChildrenHeight += totalElementHeight;

        if (articleChildrenHeight >= availableHeight) {
          breakingChildIndex = index;
          break;
        }
      }
    }

    const nodesToKeep = Array.from(article.children).slice(
      0,
      breakingChildIndex === 0 ? article.children.length : breakingChildIndex,
    );
    const nodesToMove =
      breakingChildIndex === 0 ? [] : Array.from(article.children).slice(breakingChildIndex);

    article.replaceChildren(...nodesToKeep);

    if (currentTable) {
      article.replaceChildren(currentTable, ...nodesToKeep);
    } else {
      article.replaceChildren(...nodesToKeep);
    }

    if (nodesToMove.length > 0 || newTable) {
      newSection = section.cloneNode(false) as HTMLElement;

      const newArticle = article.cloneNode(false) as HTMLElement;

      if (newTable) {
        newArticle.replaceChildren(newTable, ...nodesToMove);
      } else {
        newArticle.replaceChildren(...nodesToMove);
      }

      if (this.options.renderHeaders) {
        const header = this.renderHeaderFooter(
          sectionsProps[Number(section.id)]?.headerRefs,
          sectionsProps[Number(section.id)],
          Number(section.id) + 1,
          sectionIndex === 0,
          newSection,
        );

        if (header) {
          const parentHeaderHeight = getComputedStyles(header, "height");
          const newHeader = header.cloneNode(true) as HTMLElement;

          newHeader.style.minHeight = parentHeaderHeight;

          newSection.replaceChild(newHeader, header);
        }
      }

      newSection.appendChild(newArticle);

      // Footnotes will be rendered in the main rendering logic where the sup elements are located

      if (this.options.renderFooters) {
        const footer = this.renderHeaderFooter(
          sectionsProps[Number(section.id)]?.footerRefs,
          sectionsProps[Number(section.id)],
          Number(section.id) + 1,
          sectionIndex === 0,
          newSection,
        );

        if (footer) {
          const parentFooterHeight = getComputedStyles(footer, "height");

          const newFooter = footer.cloneNode(true) as HTMLElement;

          newFooter.style.minHeight = parentFooterHeight;

          newSection.replaceChild(newFooter, footer);
        }
      }

      /**  Since some elements didn't fit into the current section, we're transferring them to the next one. */
      replaceSectionIndex = sectionIndex + 1;
    }

    if (newSection) {
      const oldWrapper = bodyContainer.querySelector(`.${this.className}-wrapper`);

      if (oldWrapper) {
        bodyContainer.removeChild(oldWrapper);
      }

      const newSectionElements = [
        ...sections.slice(0, replaceSectionIndex),
        newSection,
        ...sections.slice(replaceSectionIndex),
      ];

      if (hasOversizeElement) {
        const modifiedSection = newSectionElements[sectionIndex].cloneNode(true) as HTMLElement;

        modifiedSection.style.minHeight = modifiedSection.style.height;
        modifiedSection.style.height = "";

        newSectionElements[sectionIndex] = modifiedSection;
      }

      const filteredNewSections = newSectionElements.filter((newSection) => {
        const article = Array.from(newSection.children).find((el) => el.localName === "article");

        return !(
          article.childNodes.length <= 1 && article.childNodes[0]?.textContent?.trim() === ""
        );
      });

      const newWrapper = this.renderWrapper(filteredNewSections);

      bodyContainer.appendChild(newWrapper);

      this.renderChildSections(filteredNewSections, bodyContainer, sectionsProps);
    }
  }

  renderChildSections(
    sectionElements: HTMLElement[],
    bodyContainer: HTMLElement,
    sectionsProps: SectionProperties[],
  ) {
    sectionElements.forEach((section, sectionIndex, sections) => {
      const article = Array.from(section.children).find((item) => item.localName === "article");

      if (!article) {
        return;
      }

      const otherChildren = Array.from(section.children).filter(
        (item) => item !== article && item.nodeType === Node.ELEMENT_NODE,
      );

      let otherChildrenHeight = 0;
      otherChildren.forEach((child) => {
        otherChildrenHeight += pxToPt(getComputedStyles(child, "height"));
      });

      const sectionHeight = pxToPt(getComputedStyles(section, "height"));
      const sectionPaddingTop = pxToPt(getComputedStyles(section, "paddingTop"));
      const sectionPaddingBottom = pxToPt(getComputedStyles(section, "paddingBottom"));

      const availableHeight =
        sectionHeight - (sectionPaddingTop + sectionPaddingBottom + otherChildrenHeight);

      const articleHeight = pxToPt(getComputedStyles(article, "height"));

      if (articleHeight <= availableHeight) {
        return;
      }

      this.getBreaksSections({
        availableHeight,
        article,
        section,
        sections,
        sectionIndex,
        bodyContainer,
        sectionsProps,
      });
    });

    // After breaking sections, render footnotes for each section
    this.renderFootnotesForSections(sectionElements);
  }

  renderFootnotesForSections(sectionElements: HTMLElement[]) {
    if (!this.options.renderFootnotes) return;

    sectionElements.forEach((section, sectionIndex) => {
      const article = Array.from(section.children).find((item) => item.localName === "article");
      if (!article) return;

      // Find all footnote references in this section
      const footnoteRefs = article.querySelectorAll("sup[id^='footnote-']");
      if (footnoteRefs.length === 0) return;

      // Extract footnote IDs
      const footnoteIds = Array.from(footnoteRefs).map((ref) => ref.id.replace("footnote-", ""));

      // Remove existing footnotes if any
      const existingFootnotes = section.querySelector("#footnote");
      if (existingFootnotes) {
        existingFootnotes.remove();
      }

      // Render footnotes for this section
      this.renderNotes(footnoteIds, this.footnoteMap, section);

      // Move footnotes before any existing footer to ensure correct order
      const renderedFootnotes = section.querySelector("#footnote");
      const existingFooter = section.querySelector("footer");

      if (renderedFootnotes && existingFooter) {
        // Move footnotes before footer
        section.insertBefore(renderedFootnotes, existingFooter);
      }
    });
  }

  renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
    const variables = {} as Record<string, any>;
    const fontScheme = themePart.theme?.fontScheme;

    if (fontScheme) {
      if (fontScheme.majorFont) {
        variables["--docx-majorHAnsi-font"] = fontScheme.majorFont.latinTypeface;
      }

      if (fontScheme.minorFont) {
        variables["--docx-minorHAnsi-font"] = fontScheme.minorFont.latinTypeface;
      }
    }

    const colorScheme = themePart.theme?.colorScheme;

    if (colorScheme) {
      for (const [key, value] of Object.entries(colorScheme.colors)) {
        variables[`--docx-${key}-color`] = `#${value}`;
      }
    }

    const cssText = this.styleToString(`.${this.className}`, variables);
    styleContainer.appendChild(this.createStyleElement(cssText));
  }

  renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
    for (const font of fontsPart.fonts) {
      for (const ref of font.embedFontRefs) {
        this.tasks.push(
          this.document.loadFont(ref.id, ref.key).then((fontData) => {
            const cssValues: Record<string, any> = {
              "font-family": encloseFontFamily(font.name),
              src: `url(${fontData})`,
            };

            if (ref.type == "bold" || ref.type == "boldItalic") {
              cssValues["font-weight"] = "bold";
            }

            if (ref.type == "italic" || ref.type == "boldItalic") {
              cssValues["font-style"] = "italic";
            }

            const cssText = this.styleToString("@font-face", cssValues);
            styleContainer.appendChild(this.createComment(`docxjs ${font.name} font`));
            styleContainer.appendChild(this.createStyleElement(cssText));
          }),
        );
      }
    }
  }

  processStyleName(className: string): string {
    return className ? `${this.className}_${escapeClassName(className)}` : this.className;
  }

  processStyles(styles: IDomStyle[]) {
    const stylesMap = keyBy(
      styles.filter((style) => style.id != null),
      (item) => item.id,
    );

    for (const style of styles.filter((style) => style.basedOn)) {
      const baseStyle = stylesMap[style.basedOn];

      if (baseStyle) {
        style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
        style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

        for (const baseValues of baseStyle.styles) {
          const styleValues = style.styles.find((style) => style.target == baseValues.target);

          if (styleValues) {
            this.copyStyleProperties(baseValues.values, styleValues.values);
          } else {
            style.styles.push({ ...baseValues, values: { ...baseValues.values } });
          }
        }
      } else if (this.options.debug) {
        // eslint-disable-next-line no-console
        console.warn(`Can't find base style ${style.basedOn}`);
      }
    }

    for (const style of styles) {
      style.cssName = this.processStyleName(style.id);
    }

    return stylesMap;
  }

  prodessNumberings(numberings: IDomNumbering[]) {
    for (const num of numberings.filter((num) => num.pStyleName)) {
      const style = this.findStyle(num.pStyleName);

      if (style?.paragraphProps?.numbering) {
        style.paragraphProps.numbering.level = num.level;
      }
    }
  }

  processElement(element: OpenXmlElement) {
    if (element.children) {
      for (const el of element.children) {
        el.parent = element;

        if (el.type == DomType.Table) {
          this.processTable(el);
        } else {
          this.processElement(el);
        }
      }
    }
  }

  processTable(table: WmlTable) {
    for (const rows of table.children) {
      for (const child of rows.children) {
        child.cssStyle = this.copyStyleProperties(table.cellStyle, child.cssStyle, [
          "border-left",
          "border-right",
          "border-top",
          "border-bottom",
          "padding-left",
          "padding-right",
          "padding-top",
          "padding-bottom",
        ]);

        this.processElement(child);
      }
    }
  }

  copyStyleProperties(
    input: Record<string, string>,
    output: Record<string, string>,
    attrs: string[] | null = null,
  ): Record<string, string> {
    if (!input) return output;

    if (output == null) output = {};
    if (attrs == null) attrs = Object.getOwnPropertyNames(input);

    for (const key of attrs) {
      if (
        Object.prototype.hasOwnProperty.call(input, key) &&
        !Object.prototype.hasOwnProperty.call(output, key)
      )
        output[key] = input[key];
    }

    return output;
  }

  createPageElement(className: string, props: SectionProperties, index: number): HTMLElement {
    const elem = this.createElement("section", { className });

    if (props) {
      elem.id = `${index}`;

      if (props.pageMargins) {
        elem.style.paddingLeft = props.pageMargins.left;
        elem.style.paddingRight = props.pageMargins.right;
        elem.style.paddingTop = props.pageMargins.top;
        elem.style.paddingBottom = props.pageMargins.bottom;
      }

      if (props.pageSize) {
        if (!this.options.ignoreWidth) elem.style.width = props.pageSize.width;

        if (!this.options.ignoreHeight && !this.options.breakPages)
          elem.style.minHeight = props.pageSize.height;

        if (this.options.breakPages) elem.style.height = props.pageSize.height;
      }
    }

    return elem;
  }

  createSectionContent(props: SectionProperties): HTMLElement {
    const elem = this.createElement("article");

    if (props.columns && props.columns.numberOfColumns) {
      elem.style.columnCount = `${props.columns.numberOfColumns}`;
      elem.style.columnGap = props.columns.space;

      if (props.columns.separator) {
        elem.style.columnRule = "1px solid black";
      }
    }

    return elem;
  }

  renderSections(document: DocumentElement) {
    const result = [];

    this.processElement(document);

    const sections =
      document.children?.length > 0 ? this.splitBySection(document.children, document.props) : [];

    const pages = this.groupByPageBreaks(sections);

    let prevProps = null;

    const sectionsProps = [];

    for (let index = 0, length = pages.length; index < length; index++) {
      this.currentFootnoteIds = [];

      const section = pages[index][0];

      let props = section.sectProps;
      const pageElement = this.createPageElement(this.className, props, index);
      this.renderStyleValues(document.cssStyle, pageElement);

      if (this.options.renderHeaders) {
        this.renderHeaderFooter(
          props.headerRefs,
          props,
          result.length,
          prevProps != props,
          pageElement,
        );
      }

      for (const sect of pages[index]) {
        const contentElement = this.createSectionContent(sect.sectProps);
        this.renderElements(sect.elements, contentElement);
        pageElement.appendChild(contentElement);

        props = sect.sectProps;
      }

      if (this.options.renderEndnotes && index == length - 1) {
        this.renderNotes(this.currentEndnoteIds, this.endnoteMap, pageElement);
      }

      result.push(pageElement);
      prevProps = props;

      sectionsProps.push(props);
    }

    return { result, sectionsProps };
  }

  renderHeaderFooter(
    refs: FooterHeaderReference[],
    props: SectionProperties,
    page: number,
    firstOfSection: boolean,
    into: HTMLElement,
  ) {
    if (!refs) return;

    const ref =
      (props.titlePage && firstOfSection ? refs.find((item) => item.type == "first") : null) ??
      (page % 2 == 1 ? refs.find((item) => item.type == "even") : null) ??
      refs.find((item) => item.type == "default");

    const part =
      ref &&
      (this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart);

    if (part) {
      this.currentPart = part;
      if (!this.usedHeaderFooterParts.includes(part.path)) {
        this.processElement(part.rootElement);
        this.usedHeaderFooterParts.push(part.path);
      }

      const [el] = this.renderElements([part.rootElement], into) as HTMLElement[];

      if (props?.pageMargins) {
        if (part.rootElement.type === DomType.Header) {
          el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
          el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
        } else if (part.rootElement.type === DomType.Footer) {
          el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
          el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
        }
      }

      this.currentPart = null;

      return el;
    }
  }

  isPageBreakElement(elem: OpenXmlElement): boolean {
    if (elem.type != DomType.Break) {
      return false;
    }

    if ((elem as WmlBreak).break == "lastRenderedPageBreak")
      return !this.options.ignoreLastRenderedPageBreak;

    return (elem as WmlBreak).break == "page";
  }

  isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean {
    if (!prev) return false;
    if (!next) return false;

    return (
      prev.pageSize?.orientation != next.pageSize?.orientation ||
      prev.pageSize?.width != next.pageSize?.width ||
      prev.pageSize?.height != next.pageSize?.height
    );
  }

  splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): Section[] {
    let current: Section = {
      sectProps: null,
      elements: [],
      pageBreak: false,
    } as unknown as Section;
    const result = [current];

    for (const elem of elements) {
      if (elem.type == DomType.Paragraph) {
        const style = this.findStyle((elem as WmlParagraph).styleName);

        if (style?.paragraphProps?.pageBreakBefore) {
          current.pageBreak = true;
          current = { sectProps: null, elements: [], pageBreak: false } as unknown as Section;
          result.push(current);
        }
      }

      current.elements.push(elem);

      if (elem.type == DomType.Paragraph) {
        const paragraph = elem as WmlParagraph;

        const sectProps = paragraph.sectionProps;
        let pBreakIndex = -1;
        let rBreakIndex = -1;

        if (this.options.breakPages && paragraph.children) {
          pBreakIndex = paragraph.children.findIndex((item) => {
            rBreakIndex = item.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
            return rBreakIndex != -1;
          });
        }

        if (sectProps || pBreakIndex != -1) {
          current.sectProps = sectProps;
          current.pageBreak = pBreakIndex != -1;
          current = { sectProps: null, elements: [], pageBreak: false } as unknown as Section;
          result.push(current);
        }

        if (pBreakIndex != -1 && paragraph.children) {
          const breakRun = paragraph.children[pBreakIndex];
          const splitRun = rBreakIndex < breakRun.children.length - 1;

          if (pBreakIndex < paragraph.children.length - 1 || splitRun) {
            const children = elem.children;
            const newParagraph = { ...elem, children: children.slice(pBreakIndex) };
            elem.children = children.slice(0, pBreakIndex);
            current.elements.push(newParagraph);

            if (splitRun) {
              const runChildren = breakRun.children;
              const newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
              elem.children.push(newRun);
              breakRun.children = runChildren.slice(rBreakIndex);
            }
          }
        }
      }
    }

    let currentSectProps = null;

    for (let index = result.length - 1; index >= 0; index--) {
      if (result[index].sectProps == null) {
        result[index].sectProps = currentSectProps ?? defaultProps;
      } else {
        currentSectProps = result[index].sectProps;
      }
    }

    /** Removing the last item in the section if this is run element with empty children  */
    for (const section of result) {
      const lastChildRunLength = section.elements[section.elements.length - 1]?.children?.find(
        (child) => child.type === "run",
      )?.children?.length;

      if (lastChildRunLength === 0) {
        section.elements = [...section.elements.slice(0, section.elements.length - 1)];
      }
    }

    return result;
  }

  groupByPageBreaks(sections: Section[]): Section[][] {
    let current: Section[] = [];
    let prev: SectionProperties;
    const result: Section[][] = [current];

    for (const section of sections) {
      current.push(section);

      if (
        this.options.ignoreLastRenderedPageBreak ||
        section.pageBreak ||
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        this.isPageBreakSection(prev, section.sectProps)
      )
        result.push((current = []));

      prev = section.sectProps;
    }

    return result.filter((item) => item.length > 0);
  }

  renderWrapper(children: HTMLElement[]) {
    return this.createElement("div", { className: `${this.className}-wrapper` }, children);
  }

  renderDefaultStyle() {
    const className = this.className;
    let wrapperStyle = `
.${className}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${className}-wrapper>section.${className} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
    if (this.options.hideWrapperOnPrint) {
      wrapperStyle = `@media not print { ${wrapperStyle} }`;
    }
    let styleText = `${wrapperStyle}
.${className} { color: black; hyphens: auto; text-underline-position: from-font; }
section.${className} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${className}>article { margin-bottom: auto; z-index: 1; }
section.${className}>footer { z-index: 1; }
.${className} table { border-collapse: collapse; }
.${className} table td, .${className} table th { vertical-align: top; }
.${className} p { margin: 0pt; min-height: 1em; }
.${className} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${className} a { color: inherit; text-decoration: inherit; }
.${className} svg { fill: transparent; }
`;

    if (this.options.renderComments) {
      styleText += `
.${className}-comment-ref { cursor: default; }
.${className}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${className}-comment-ref:hover~.${className}-comment-popover { display: block; }
.${className}-comment-author,.${className}-comment-date { font-size: 0.875rem; color: #888; }
`;
    }

    return this.createStyleElement(styleText);
  }

  renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
    let styleText = "";
    const resetCounters = [];

    for (const num of numberings) {
      const selector = `p.${this.numberingClass(num.id, num.level)}`;
      let listStyleType = "none";

      if (num.bullet) {
        const valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

        styleText += this.styleToString(
          `${selector}:before`,
          {
            content: "' '",
            display: "inline-block",
            background: `var(${valiable})`,
          },
          num.bullet.style,
        );

        this.tasks.push(
          this.document.loadNumberingImage(num.bullet.src).then((data) => {
            const text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
            styleContainer.appendChild(this.createStyleElement(text));
          }),
        );
      } else if (num.levelText) {
        const counter = this.numberingCounter(num.id, num.level);
        const counterReset = counter + " " + (num.start - 1);

        if (num.level > 0) {
          styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
            "counter-set": counterReset,
          });
        }
        // reset all level counters with start value
        resetCounters.push(counterReset);

        styleText += this.styleToString(`${selector}:before`, {
          content: this.levelTextToContent(
            num.levelText,
            num.suff,
            num.id,
            this.numFormatToCssValue(num.format),
          ),
          "counter-increment": counter,
          ...num.rStyle,
        });
      } else {
        listStyleType = this.numFormatToCssValue(num.format);
      }

      styleText += this.styleToString(selector, {
        display: "list-item",
        "list-style-position": "inside",
        "list-style-type": listStyleType,
        ...num.pStyle,
      });
    }

    if (resetCounters.length > 0) {
      styleText += this.styleToString(this.rootSelector, {
        "counter-reset": resetCounters.join(" "),
      });
    }

    return this.createStyleElement(styleText);
  }

  renderStyles(styles: IDomStyle[]): HTMLElement {
    let styleText = "";
    const stylesMap = this.styleMap;
    const defautStyles = keyBy(
      styles.filter((style) => style.isDefault),
      (style) => style.target,
    );

    for (const style of styles) {
      let subStyles = style.styles;

      if (style.linked) {
        const linkedStyle = style.linked && stylesMap[style.linked];

        if (linkedStyle) {
          subStyles = subStyles.concat(linkedStyle.styles);
        } else if (this.options.debug) {
          // eslint-disable-next-line no-console
          console.warn(`Can't find linked style ${style.linked}`);
        }
      }

      for (const subStyle of subStyles) {
        //TODO temporary disable modificators until test it well
        let selector = `${style.target ?? ""}.${style.cssName}`; //${subStyle.mod ?? ''}

        if (style.target != subStyle.target) selector += ` ${subStyle.target}`;

        if (defautStyles[style.target] == style)
          selector = `.${this.className} ${style.target}, ` + selector;

        styleText += this.styleToString(selector, subStyle.values);
      }
    }

    return this.createStyleElement(styleText);
  }

  renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, into: HTMLElement) {
    const notes = noteIds.map((id) => notesMap[id]).filter((note) => note);

    if (notes.length > 0) {
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      const result = this.createElement("ol", null, this.renderElements(notes));
      result.id = "footnote";

      into.appendChild(result);
    }
  }

  renderElement(elem: OpenXmlElement): Node | Node[] | null {
    switch (elem.type) {
      case DomType.Paragraph:
        return this.renderParagraph(elem as WmlParagraph);

      case DomType.BookmarkStart:
        return this.renderBookmarkStart(elem as WmlBookmarkStart);

      case DomType.BookmarkEnd:
        return null; //ignore bookmark end

      case DomType.Run:
        return this.renderRun(elem as WmlRun);

      case DomType.Table:
        return this.renderTable(elem);

      case DomType.Row:
        return this.renderTableRow(elem);

      case DomType.Cell:
        return this.renderTableCell(elem);

      case DomType.Hyperlink:
        return this.renderHyperlink(elem);

      case DomType.SmartTag:
        return this.renderSmartTag(elem);

      case DomType.Drawing:
        return this.renderDrawing(elem);

      case DomType.Image:
        return this.renderImage(elem as IDomImage);

      case DomType.Text:
        return this.renderText(elem as WmlText);

      case DomType.DeletedText:
        return this.renderDeletedText(elem as WmlText);

      case DomType.Tab:
        return this.renderTab(elem);

      case DomType.Symbol:
        return this.renderSymbol(elem as WmlSymbol);

      case DomType.Break:
        return this.renderBreak(elem as WmlBreak);

      case DomType.Footer:
        return this.renderContainer(elem, "footer");

      case DomType.Header:
        return this.renderContainer(elem, "header");

      case DomType.Footnote:
      case DomType.Endnote:
        return this.renderContainer(elem, "li");

      case DomType.FootnoteReference:
        return this.renderFootnoteReference(elem as WmlNoteReference);

      case DomType.EndnoteReference:
        return this.renderEndnoteReference(elem as WmlNoteReference);

      case DomType.NoBreakHyphen:
        return this.createElement("wbr");

      case DomType.VmlPicture:
        return this.renderVmlPicture(elem);

      case DomType.VmlElement:
        return this.renderVmlElement(elem as VmlElement);

      case DomType.MmlMath:
        return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });

      case DomType.MmlMathParagraph:
        return this.renderContainer(elem, "span");

      case DomType.MmlFraction:
        return this.renderContainerNS(elem, ns.mathML, "mfrac");

      case DomType.MmlBase:
        return this.renderContainerNS(
          elem,
          ns.mathML,
          elem?.parent?.type == DomType.MmlMatrixRow ? "mtd" : "mrow",
        );

      case DomType.MmlNumerator:
      case DomType.MmlDenominator:
      case DomType.MmlFunction:
      case DomType.MmlLimit:
      case DomType.MmlBox:
        return this.renderContainerNS(elem, ns.mathML, "mrow");

      case DomType.MmlGroupChar:
        return this.renderMmlGroupChar(elem);

      case DomType.MmlLimitLower:
        return this.renderContainerNS(elem, ns.mathML, "munder");

      case DomType.MmlMatrix:
        return this.renderContainerNS(elem, ns.mathML, "mtable");

      case DomType.MmlMatrixRow:
        return this.renderContainerNS(elem, ns.mathML, "mtr");

      case DomType.MmlRadical:
        return this.renderMmlRadical(elem);

      case DomType.MmlSuperscript:
        return this.renderContainerNS(elem, ns.mathML, "msup");

      case DomType.MmlSubscript:
        return this.renderContainerNS(elem, ns.mathML, "msub");

      case DomType.MmlDegree:
      case DomType.MmlSuperArgument:
      case DomType.MmlSubArgument:
        return this.renderContainerNS(elem, ns.mathML, "mn");

      case DomType.MmlFunctionName:
        return this.renderContainerNS(elem, ns.mathML, "ms");

      case DomType.MmlDelimiter:
        return this.renderMmlDelimiter(elem);

      case DomType.MmlRun:
        return this.renderMmlRun(elem);

      case DomType.MmlNary:
        return this.renderMmlNary(elem);

      case DomType.MmlPreSubSuper:
        return this.renderMmlPreSubSuper(elem);

      case DomType.MmlBar:
        return this.renderMmlBar(elem);

      case DomType.MmlEquationArray:
        return this.renderMllList(elem);

      case DomType.Inserted:
        return this.renderInserted(elem);

      case DomType.Deleted:
        return this.renderDeleted(elem);

      case DomType.CommentRangeStart:
        return this.renderCommentRangeStart(elem);

      case DomType.CommentRangeEnd:
        return this.renderCommentRangeEnd(elem);

      case DomType.CommentReference:
        return this.renderCommentReference(elem);

      case DomType.AltChunk:
        return this.renderAltChunk(elem);
    }

    return null;
  }

  renderElements(elems: OpenXmlElement[], into?: Node): Node[] | null {
    if (elems == null) return null;

    const result = elems.flatMap((el) => this.renderElement(el)).filter((el) => el != null);

    if (into) {
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      appendChildren(into, result);
    }

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    return result;
  }

  renderContainer<T extends keyof HTMLElementTagNameMap>(
    elem: OpenXmlElement,
    tagName: T,
    props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>,
  ): HTMLElementTagNameMap[T] {
    return this.createElement<T>(
      tagName,
      props,
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      this.renderElements(elem.children),
    );
  }

  renderContainerNS(
    elem: OpenXmlElement,
    ns: string,
    tagName: string,
    props?: Record<string, any>,
  ) {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    return this.createElementNS(ns, tagName, props, this.renderElements(elem.children));
  }

  renderParagraph(elem: WmlParagraph) {
    const result = this.renderContainer(elem, "p");

    const style = this.findStyle(elem.styleName);
    elem.tabs ??= style?.paragraphProps?.tabs;

    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);
    this.renderCommonProperties(result.style, elem);

    const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

    if (numbering) {
      result.classList.add(this.numberingClass(numbering.id, numbering.level));
    }

    return result;
  }

  renderRunProperties(style: any, props: RunProperties) {
    this.renderCommonProperties(style, props);
  }

  renderCommonProperties(style: any, props: CommonProperties) {
    if (props == null) return;

    if (props.color) {
      style.color = props.color;
    }

    if (props.fontSize) {
      style["font-size"] = props.fontSize;
    }
  }

  renderHyperlink(elem: WmlHyperlink) {
    const result = this.renderContainer(elem, "a");

    this.renderStyleValues(elem.cssStyle, result);

    let href = "";

    if (elem.id) {
      const rel = this.document.documentPart.rels.find(
        (it) => it.id == elem.id && it.targetMode === "External",
      );
      href = rel?.target ?? href;
    }

    if (elem.anchor) {
      href += `#${elem.anchor}`;
    }

    result.href = href;

    return result;
  }

  renderSmartTag(elem: WmlSmartTag) {
    return this.renderContainer(elem, "span");
  }

  renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
    if (!this.options.renderComments) return null;

    const rng = new Range();
    this.commentHighlight?.add(rng);

    const result = this.htmlDocument.createComment(`start of comment #${commentStart.id}`);
    this.later(() => rng.setStart(result, 0));
    this.commentMap[commentStart.id] = rng;

    return result;
  }

  renderCommentRangeEnd(commentEnd: WmlCommentRangeStart) {
    if (!this.options.renderComments) return null;

    const rng = this.commentMap[commentEnd.id];
    const result = this.htmlDocument.createComment(`end of comment #${commentEnd.id}`);
    this.later(() => rng?.setEnd(result, 0));

    return result;
  }

  renderCommentReference(commentRef: WmlCommentReference) {
    if (!this.options.renderComments) return null;

    const comment = this.document.commentsPart?.commentMap[commentRef.id];

    if (!comment) return null;

    const frg = new DocumentFragment();
    const commentRefEl = this.createElement(
      "span",
      { className: `${this.className}-comment-ref` },
      ["??"],
    );
    const commentsContainerEl = this.createElement("div", {
      className: `${this.className}-comment-popover`,
    });

    this.renderCommentContent(comment, commentsContainerEl);

    frg.appendChild(
      this.htmlDocument.createComment(
        `comment #${comment.id} by ${comment.author} on ${comment.date}`,
      ),
    );
    frg.appendChild(commentRefEl);
    frg.appendChild(commentsContainerEl);

    return frg;
  }

  renderAltChunk(elem: WmlAltChunk) {
    if (!this.options.renderAltChunks) return null;

    const result = this.createElement("iframe");

    this.tasks.push(
      this.document.loadAltChunk(elem.id, this.currentPart).then((res) => {
        result.srcdoc = res;
      }),
    );

    return result;
  }

  renderCommentContent(comment: WmlComment, container: Node) {
    container.appendChild(
      this.createElement("div", { className: `${this.className}-comment-author` }, [
        comment.author,
      ]),
    );
    container.appendChild(
      this.createElement("div", { className: `${this.className}-comment-date` }, [
        new Date(comment.date).toLocaleString(),
      ]),
    );

    this.renderElements(comment.children, container);
  }

  renderDrawing(elem: OpenXmlElement) {
    const result = this.renderContainer(elem, "div");

    result.style.display = "inline-block";
    result.style.position = "relative";
    result.style.textIndent = "0px";

    this.renderStyleValues(elem.cssStyle, result);

    return result;
  }

  renderImage(elem: IDomImage) {
    const result = this.createElement("img");

    this.renderStyleValues(elem.cssStyle, result);

    if (this.document) {
      this.tasks.push(
        this.document.loadDocumentImage(elem.src, this.currentPart).then((res) => {
          result.src = res;
        }),
      );
    }

    return result;
  }

  renderText(elem: WmlText) {
    return this.htmlDocument.createTextNode(elem.text);
  }

  renderDeletedText(elem: WmlText) {
    return this.options.renderEndnotes ? this.htmlDocument.createTextNode(elem.text) : null;
  }

  renderBreak(elem: WmlBreak) {
    if (elem.break == "textWrapping") {
      return this.createElement("br");
    }

    return null;
  }

  renderInserted(elem: OpenXmlElement): Node | Node[] {
    if (this.options.renderChanges) return this.renderContainer(elem, "ins");

    return this.renderElements(elem.children);
  }

  renderDeleted(elem: OpenXmlElement): Node | null {
    if (this.options.renderChanges) return this.renderContainer(elem, "del");

    return null;
  }

  renderSymbol(elem: WmlSymbol) {
    const span = this.createElement("span");
    span.style.fontFamily = elem.font;
    span.innerHTML = `&#x${elem.char};`;
    return span;
  }

  renderFootnoteReference(elem: WmlNoteReference) {
    const result = this.createElement("sup");
    this.currentFootnoteIds.push(elem.id);
    result.textContent = `${this.currentFootnoteIds.length}`;
    result.id = `footnote-${elem.id}`;

    return result;
  }

  renderEndnoteReference(elem: WmlNoteReference) {
    const result = this.createElement("sup");
    this.currentEndnoteIds.push(elem.id);
    result.textContent = `${this.currentEndnoteIds.length}`;

    return result;
  }

  renderTab(elem: OpenXmlElement) {
    const tabSpan = this.createElement("span");

    tabSpan.innerHTML = "&emsp;"; //"&nbsp;";

    if (this.options.experimental) {
      tabSpan.className = this.tabStopClass();
      const stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
      this.currentTabs.push({ stops, span: tabSpan });
    }

    return tabSpan;
  }

  renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
    return this.createElement("span", { id: elem.name });
  }

  renderRun(elem: WmlRun) {
    if (elem.fieldRun) return null;

    const result = this.createElement("span");

    if (elem.id) result.id = elem.id;

    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);

    if (elem.verticalAlign) {
      const wrapper = this.createElement(elem.verticalAlign as any);
      this.renderElements(elem.children, wrapper);
      result.appendChild(wrapper);
    } else {
      this.renderElements(elem.children, result);
    }

    return result;
  }

  normalizeTableBorders(table: HTMLTableElement, deepLevel = 0) {
    const rows = table.querySelectorAll(":scope > tr");

    rows.forEach((row, rowIndex) => {
      const cells = row.querySelectorAll<HTMLTableCellElement>(":scope > td");

      cells.forEach((cell, cellIndex) => {
        const nestedTables = cell.querySelectorAll<HTMLTableElement>(":scope > table");

        nestedTables.forEach((nestedTable) => {
          this.normalizeTableBorders(nestedTable, deepLevel + 1);
        });

        if (deepLevel === 0) {
          cell.style.width = "100%";
        }

        if (deepLevel > 0) {
          cell.style.borderTop = "none";
          cell.style.borderBottom = "none";
          cell.style.borderRight = "none";
          cell.style.borderLeft = "none";
        }

        if (cellIndex > 0) {
          cell.style.borderLeft = "0.5pt solid black";
        }

        if (rows.length > 1 && rowIndex !== rows.length - 1) {
          cell.style.borderBottom = "0.5pt solid black";
        }
      });
    });
  }

  removeEmptyParagraphs(table: HTMLTableElement) {
    const parentTd = table.closest("td");
    const emptyParagraphs =
      parentTd?.querySelectorAll("p:empty") ?? table.querySelectorAll("p:empty");

    emptyParagraphs.forEach((paragraph) => {
      if (paragraph.textContent.trim() === "" && paragraph.children.length === 0) {
        paragraph.remove();
      }
    });
  }

  renderTable(elem: WmlTable) {
    const table = this.createElement("table");

    table.setAttribute("data-parsed-table", "true");

    this.tableCellPositions.push(this.currentCellPosition);
    this.tableVerticalMerges.push(this.currentVerticalMerge);
    this.currentVerticalMerge = {};
    this.currentCellPosition = { col: 0, row: 0 };

    if (elem.columns) table.appendChild(this.renderTableColumns(elem.columns));

    this.renderClass(elem, table);
    this.renderElements(elem.children, table);
    this.renderStyleValues(elem.cssStyle, table);

    this.currentVerticalMerge = this.tableVerticalMerges.pop();
    this.currentCellPosition = this.tableCellPositions.pop();

    const hasNestedTable = table.querySelector("table");

    /** if there are nested tables call function to merge borders there */
    if (hasNestedTable) {
      this.normalizeTableBorders(table);

      this.removeEmptyParagraphs(table);
    }

    return table;
  }

  renderTableColumns(columns: WmlTableColumn[]) {
    const result = this.createElement("colgroup");

    for (const col of columns) {
      const colElem = this.createElement("col");

      if (col.width) colElem.style.width = col.width;

      result.appendChild(colElem);
    }

    return result;
  }

  renderTableRow(elem: WmlTableRow) {
    const result = this.createElement("tr");

    this.currentCellPosition.col = 0;

    if (elem.gridBefore) result.appendChild(this.renderTableCellPlaceholder(elem.gridBefore));

    this.renderClass(elem, result);
    this.renderElements(elem.children, result);
    this.renderStyleValues(elem.cssStyle, result);

    if (elem.gridAfter) result.appendChild(this.renderTableCellPlaceholder(elem.gridAfter));

    this.currentCellPosition.row++;

    return result;
  }

  renderTableCellPlaceholder(colSpan: number) {
    const result = this.createElement("td", { colSpan });
    result.style.border = "none";
    return result;
  }

  renderTableCell(elem: WmlTableCell) {
    const result = this.renderContainer(elem, "td");

    const key = this.currentCellPosition.col;

    if (elem.verticalMerge) {
      if (elem.verticalMerge == "restart") {
        this.currentVerticalMerge[key] = result;
        result.rowSpan = 1;
      } else if (this.currentVerticalMerge[key]) {
        this.currentVerticalMerge[key].rowSpan += 1;

        result.remove();

        return null;
      }
    } else {
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      this.currentVerticalMerge[key] = null;
    }

    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);

    if (elem.span) result.colSpan = elem.span;

    this.currentCellPosition.col += result.colSpan;

    return result;
  }

  renderVmlPicture(elem: OpenXmlElement) {
    return this.renderContainer(elem, "div");
  }

  renderVmlElement(elem: VmlElement): SVGElement {
    const container = this.createSvgElement("svg");

    container.setAttribute("style", elem.cssStyleText);

    const result = this.renderVmlChildElement(elem);

    if (elem.imageHref?.id) {
      this.tasks.push(
        this.document
          ?.loadDocumentImage(elem.imageHref.id, this.currentPart)
          .then((res) => result.setAttribute("href", res)),
      );
    }

    container.appendChild(result);

    requestAnimationFrame(() => {
      const bb = (container.firstElementChild as any).getBBox();

      container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
      container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
    });

    return container;
  }

  renderVmlChildElement(elem: VmlElement): any {
    const result = this.createSvgElement(elem.tagName as any);
    Object.entries(elem.attrs).forEach(([key, value]) => result.setAttribute(key, value));

    for (const child of elem.children) {
      if (child.type == DomType.VmlElement) {
        result.appendChild(this.renderVmlChildElement(child as VmlElement));
      } else {
        result.appendChild(...asArray(this.renderElement(child as any)));
      }
    }

    return result;
  }

  renderMmlRadical(elem: OpenXmlElement): HTMLElement {
    const elemChild = elem.children;

    const base = elemChild.find((el) => el.type == DomType.MmlBase);

    if (elem.props?.hideDegree) {
      return this.createElementNS(
        ns.mathML,
        "msqrt",
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        null,
        this.renderElements([base]),
      );
    }

    const degree = elemChild.find((el) => el.type == DomType.MmlDegree);
    return this.createElementNS(
      ns.mathML,
      "mroot",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
      this.renderElements([base, degree]),
    );
  }

  renderMmlDelimiter(elem: OpenXmlElement): HTMLElement {
    const children = [];

    children.push(
      this.createElementNS(
        ns.mathML,
        "mo",
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        null,
        [elem.props.beginChar ?? "("],
      ),
    );
    children.push(...this.renderElements(elem.children));
    children.push(
      this.createElementNS(
        ns.mathML,
        "mo",
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        null,
        [elem.props.endChar ?? ")"],
      ),
    );

    return this.createElementNS(
      ns.mathML,
      "mrow",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
      children,
    );
  }

  renderMmlNary(elem: OpenXmlElement): HTMLElement {
    const children = [];
    const grouped = keyBy(elem.children, (el) => el.type);

    const sup = grouped[DomType.MmlSuperArgument];
    const sub = grouped[DomType.MmlSubArgument];
    const supElem = sup
      ? this.createElementNS(
          ns.mathML,
          "mo",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          asArray(this.renderElement(sup)),
        )
      : null;
    const subElem = sub
      ? this.createElementNS(
          ns.mathML,
          "mo",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          asArray(this.renderElement(sub)),
        )
      : null;

    const charElem = this.createElementNS(
      ns.mathML,
      "mo",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
      [elem.props?.char ?? "\u222B"],
    );

    if (supElem || subElem) {
      children.push(
        this.createElementNS(
          ns.mathML,
          "munderover",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          [charElem, subElem, supElem],
        ),
      );
      // eslint-disable-next-line no-dupe-else-if
    } else if (supElem) {
      children.push(
        this.createElementNS(
          ns.mathML,
          "mover",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          [charElem, supElem],
        ),
      );
      // eslint-disable-next-line no-dupe-else-if
    } else if (subElem) {
      children.push(
        this.createElementNS(
          ns.mathML,
          "munder",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          [charElem, subElem],
        ),
      );
    } else {
      children.push(charElem);
    }

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    children.push(...this.renderElements(grouped[DomType.MmlBase].children));

    return this.createElementNS(
      ns.mathML,
      "mrow",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
      children,
    );
  }

  renderMmlPreSubSuper(elem: OpenXmlElement) {
    const children = [];
    const grouped = keyBy(elem.children, (el) => el.type);

    const sup = grouped[DomType.MmlSuperArgument];
    const sub = grouped[DomType.MmlSubArgument];
    const supElem = sup
      ? this.createElementNS(
          ns.mathML,
          "mo",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          asArray(this.renderElement(sup)),
        )
      : null;
    const subElem = sub
      ? this.createElementNS(
          ns.mathML,
          "mo",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          asArray(this.renderElement(sub)),
        )
      : null;
    const stubElem = this.createElementNS(
      ns.mathML,
      "mo",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
    );

    children.push(
      this.createElementNS(
        ns.mathML,
        "msubsup",
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        null,
        [stubElem, subElem, supElem],
      ),
    );

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    children.push(...this.renderElements(grouped[DomType.MmlBase].children));

    return this.createElementNS(
      ns.mathML,
      "mrow",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
      children,
    );
  }

  renderMmlGroupChar(elem: OpenXmlElement) {
    const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
    const result = this.renderContainerNS(elem, ns.mathML, tagName);

    if (elem.props.char) {
      result.appendChild(
        this.createElementNS(
          ns.mathML,
          "mo",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          [elem.props.char],
        ),
      );
    }

    return result;
  }

  renderMmlBar(elem: OpenXmlElement) {
    const result = this.renderContainerNS(elem, ns.mathML, "mrow");

    switch (elem.props.position) {
      case "top":
        result.style.textDecoration = "overline";
        break;
      case "bottom":
        result.style.textDecoration = "underline";
        break;
    }

    return result;
  }

  renderMmlRun(elem: OpenXmlElement) {
    const result = this.createElementNS(
      ns.mathML,
      "ms",
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      null,
      this.renderElements(elem.children),
    );

    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);

    return result;
  }

  renderMllList(elem: OpenXmlElement) {
    const result = this.createElementNS(ns.mathML, "mtable");

    this.renderClass(elem, result);
    this.renderStyleValues(elem.cssStyle, result);

    for (const child of this.renderElements(elem.children)) {
      result.appendChild(
        this.createElementNS(
          ns.mathML,
          "mtr",
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          null,
          [
            this.createElementNS(
              ns.mathML,
              "mtd",
              // eslint-disable-next-line @typescript-eslint/ban-ts-comment
              // @ts-ignore
              null,
              [child],
            ),
          ],
        ),
      );
    }

    return result;
  }

  renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
    for (const key in style) {
      if (key.startsWith("$")) {
        ouput.setAttribute(key.slice(1), style[key]);
      } else {
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        ouput.style[key] = style[key];
      }
    }
  }

  renderClass(input: OpenXmlElement, ouput: HTMLElement) {
    if (input.className) ouput.className = input.className;

    if (input.styleName) ouput.classList.add(this.processStyleName(input.styleName));
  }

  findStyle(styleName: string) {
    return styleName && this.styleMap?.[styleName];
  }

  numberingClass(id: string, lvl: number) {
    return `${this.className}-num-${id}-${lvl}`;
  }

  tabStopClass() {
    return `${this.className}-tab-stop`;
  }

  styleToString(selectors: string, values: Record<string, string>, cssText: string | null = null) {
    let result = `${selectors} {\r\n`;

    for (const key in values) {
      if (key.startsWith("$")) continue;

      result += `  ${key}: ${values[key]};\r\n`;
    }

    if (cssText) result += cssText;

    return result + "}\r\n";
  }

  numberingCounter(id: string, lvl: number) {
    return `${this.className}-num-${id}-${lvl}`;
  }

  levelTextToContent(text: string, suff: string, id: string, numformat: string) {
    const suffMap = {
      tab: "\\9",
      space: "\\a0",
    };

    const result = text.replace(/%\d*/g, (item) => {
      const lvl = parseInt(item.substring(1), 10) - 1;
      return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
    });

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    return `"${result}${suffMap[suff] ?? ""}"`;
  }

  numFormatToCssValue(format: string) {
    const mapping = {
      none: "none",
      bullet: "disc",
      decimal: "decimal",
      lowerLetter: "lower-alpha",
      upperLetter: "upper-alpha",
      lowerRoman: "lower-roman",
      upperRoman: "upper-roman",
      decimalZero: "decimal-leading-zero", // 01,02,03,...
      // ordinal: "", // 1st, 2nd, 3rd,...
      // ordinalText: "", //First, Second, Third, ...
      // cardinalText: "", //One,Two Three,...
      // numberInDash: "", //-1-,-2-,-3-, ...
      // hex: "upper-hexadecimal",
      aiueo: "katakana",
      aiueoFullWidth: "katakana",
      chineseCounting: "simp-chinese-informal",
      chineseCountingThousand: "simp-chinese-informal",
      chineseLegalSimplified: "simp-chinese-formal", // ????
      chosung: "hangul-consonant",
      ideographDigital: "cjk-ideographic",
      ideographTraditional: "cjk-heavenly-stem", // ???
      ideographLegalTraditional: "trad-chinese-formal",
      ideographZodiac: "cjk-earthly-branch", // ????
      iroha: "katakana-iroha",
      irohaFullWidth: "katakana-iroha",
      japaneseCounting: "japanese-informal",
      japaneseDigitalTenThousand: "cjk-decimal",
      japaneseLegal: "japanese-formal",
      thaiNumbers: "thai",
      koreanCounting: "korean-hangul-formal",
      koreanDigital: "korean-hangul-formal",
      koreanDigital2: "korean-hanja-informal",
      hebrew1: "hebrew",
      hebrew2: "hebrew",
      hindiNumbers: "devanagari",
      ganada: "hangul",
      taiwaneseCounting: "cjk-ideographic",
      taiwaneseCountingThousand: "cjk-ideographic",
      taiwaneseDigital: "cjk-decimal",
    };

    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    return mapping[format] ?? format;
  }

  refreshTabStops() {
    if (!this.options.experimental) return;

    setTimeout(() => {
      const pixelToPoint = computePixelToPoint();

      for (const tab of this.currentTabs) {
        updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
      }
    }, 500);
  }

  createElementNS(
    ns: string,
    tagName: string,
    props?: Partial<Record<any, any>>,
    children?: ChildType[],
  ): any {
    const result = ns
      ? this.htmlDocument.createElementNS(ns, tagName)
      : this.htmlDocument.createElement(tagName);
    Object.assign(result, props);

    if (children) {
      appendChildren(result, children);
    }

    return result;
  }

  createElement<T extends keyof HTMLElementTagNameMap>(
    tagName: T,
    props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>,
    children?: ChildType[],
  ): HTMLElementTagNameMap[T] {
    return this.createElementNS(
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      undefined,
      tagName,
      props,
      children,
    );
  }

  createSvgElement<T extends keyof SVGElementTagNameMap>(
    tagName: T,
    props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>,
    children?: ChildType[],
  ): SVGElementTagNameMap[T] {
    return this.createElementNS(ns.svg, tagName, props, children);
  }

  createStyleElement(cssText: string) {
    return this.createElement("style", { innerHTML: cssText });
  }

  createComment(text: string) {
    return this.htmlDocument.createComment(text);
  }

  later(func: any) {
    this.postRenderTasks.push(func);
  }
}

type ChildType = Node | string;

function removeAllElements(elem: HTMLElement) {
  elem.innerHTML = "";
}

function appendChildren(elem: Node, children: (Node | string)[]) {
  children.forEach((child) =>
    elem.appendChild(isString(child) ? document.createTextNode(child) : child),
  );
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
  let parent = elem.parent;

  while (parent != null && parent.type != type) parent = parent.parent;

  return parent as T;
}

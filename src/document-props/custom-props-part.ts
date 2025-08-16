import { Part } from "../common";
import type { CustomProperty } from "./custom-props";
import { parseCustomProps } from "./custom-props";

export class CustomPropsPart extends Part {
  props: CustomProperty[];

  parseXml(root: Element) {
    this.props = parseCustomProps(root, this._package.xmlParser);
  }
}

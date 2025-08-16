import { Part } from "../common/part";
import type { CorePropsDeclaration} from "./core-props";
import { parseCoreProps } from "./core-props";

export class CorePropsPart extends Part {
  props: CorePropsDeclaration;

  parseXml(root: Element) {
    this.props = parseCoreProps(root, this._package.xmlParser);
  }
}

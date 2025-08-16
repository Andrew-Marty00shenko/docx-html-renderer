import type { OpenXmlPackage } from "../common";
import { Part } from "../common";
import type { WmlSettings } from "./settings";
import { parseSettings } from "./settings";

export class SettingsPart extends Part {
  settings: WmlSettings;

  constructor(pkg: OpenXmlPackage, path: string) {
    super(pkg, path);
  }

  parseXml(root: Element) {
    this.settings = parseSettings(root, this._package.xmlParser);
  }
}

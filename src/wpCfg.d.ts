import { WebParts } from "gd-sprest-bs";

/**
 * Bootstrap WebPart Demo Configuration
 */
export interface IBSDemoCfg extends WebParts.IWPListCfg {
    FieldDisplayNames: Array<string>;
    FieldInternalNames: Array<string>;
}
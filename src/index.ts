import { Helper } from "gd-sprest";
import { Configuration } from "./cfg";
import { WebPart } from "./wp";

// Create the global variable
window["BSDemo"] = {
    Configuration,
    WebPart
};

// Notify the SharePoint Script-on-Demand (SOD) Library
window["SP"] ? window["SP"].SOD.notifyScriptLoadedAndExecuteWaitingJobs("bsDemo") : null;

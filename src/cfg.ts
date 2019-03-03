import { Helper } from "gd-sprest";

/**
 * Configuration
 */
export const Configuration = {
    // Make the script globally available in the site collection
    CustomAction: Helper.SPConfig({
        CustomActionCfg: {
            Site: [
                {
                    Name: "bsDemo",
                    Title: "Bootstrap Demo",
                    Description: "Example webpart for a bootstrap webpart.",
                    Location: "ScriptLink",
                    ScriptBlock: "SP.SOD.registerSod('bsDemo', '/sites/demo/siteassets/bsDemo.js');"
                }
            ]
        }
    }),

    // Make the solution available through the webpart gallery
    WebPart: Helper.SPConfig({
        WebPartCfg: [
            {
                FileName: "demo-bootstrap-list.webpart",
                Group: "Demo",
                XML: `<?xml version="1.0" encoding="utf-8"?>
<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
            <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
            <importErrorMessage>$Resources:core,ImportantErrorMessage;</importErrorMessage>
        </metaData>
        <data>
            <properties>
                <property name="Title" type="string">Bootstrap List</property>
                <property name="Description" type="string">Displays a list in a bootstrap table.</property>
                <property name="ChromeType" type="chrometype">TitleOnly</property>
                <property name="Content" type="string">
                    &lt;div id="wp-bsDemo"&gt;&lt;/div&gt;
                    &lt;div id="wp-bsDemoCfg" style="display: none;"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript"&gt;if(SP.SOD.executeOrDelayUntilScriptLoaded(function() { new BSDemo.WebPart(); }, 'bsDemo') == false) { LoadSodByKey("bsDemo"); }&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`
            }
        ]
    })
};
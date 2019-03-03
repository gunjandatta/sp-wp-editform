import { SPTypes } from "gd-sprest";
import { Components, WebParts } from "gd-sprest-bs";
import { editForm, IBSDemoCfg } from "./wpCfg";

/**
 * WebPart
 */
export const WebPart = () => {
    // Method to get the columns
    let getColumns = (cfg: IBSDemoCfg) => {
        // Set the default fields
        let columns: Array<Components.ITableColumn> = [
            // The Item ID column will hold the action button
            {
                name: "ID",
                title: "Item ID",
                isHidden: true,
                onRenderCell: (el, col, item) => {
                    // Clear the element
                    while (el.firstChild) { el.removeChild(el.firstChild); }

                    // Render an edit button
                    Components.Button({
                        el,
                        data: item,
                        type: Components.ButtonTypes.Info,
                        text: "Edit",
                        onClick: () => {
                            // Get the dialog element
                            let elDialog = document.querySelector("#dlg-bs-demo");
                            if (elDialog == null) {
                                // Create the element
                                elDialog = document.createElement("dlg-bs-demo");

                                // Add it to the body
                                document.body.appendChild(elDialog);
                            }

                            // Display an edit form dialog
                            Components.ListFormDialog({
                                el: elDialog,
                                listName: cfg.ListName,
                                itemId: item.ID,
                                controlMode: SPTypes.ControlMode.Edit,
                                visible: true,
                                modalProps: {
                                    onClose: () => {
                                        // Remove the dialog element
                                        document.body.removeChild(elDialog);
                                    }
                                }
                            });
                        }
                    });
                }
            }
        ];

        // See if custom fields have been defined
        if (cfg.FieldDisplayNames && cfg.FieldDisplayNames.length == 0) {
            // Return the title field
            columns.push({ name: "Title", title: "Title" });
        } else {
            // Parse the fields
            for (let i = 0; i < cfg.FieldDisplayNames.length; i++) {
                // Add the column
                columns.push({
                    name: cfg.FieldInternalNames[i],
                    title: cfg.FieldDisplayNames[i]
                });
            }
        }

        // Return the columns
        return columns;
    }

    // Return a list webpart
    return WebParts.WPList({
        cfgElementId: "wp-bsDemoCfg",
        elementId: "wp-bsDemo",
        editForm,
        onRenderItems: (wpInfo, items) => {
            // Clear the element
            while (wpInfo.el.firstChild) { wpInfo.el.removeChild(wpInfo.el.firstChild); }

            // See if items exist
            if (items.length == 0) {
                // Render a message
                wpInfo.el.innerHTML = "<p>No items exist in the list.</p>";
            } else {
                // Render a table
                Components.Table({
                    el: wpInfo.el,
                    rows: items,
                    columns: getColumns(wpInfo.cfg as IBSDemoCfg)
                });

            }
        }
    });
}
import { Types } from "gd-sprest";
import { Components, WebParts } from "gd-sprest-bs";
import { IBSDemoCfg } from "./wpCfg.d"
export { IBSDemoCfg }

/**
 * WebPart Edit Form
 */
export const editForm: WebParts.IWPListEditForm = ((): WebParts.IWPListEditForm => {
    let _ddl: Components.IDropdown = null;

    // Method to initialize the list fields dropdown
    let initFields = (control: Components.IFormControlPropsDropdown, list: Types.SP.List) => {
        // Return a promise, to delay the rendering until the fields have been loaded
        return new Promise((resolve, reject) => {
            // Ensure a list exists
            if (list) {
                // Load the list fields
                list.Fields().execute(fields => {
                    // Parse the fields
                    for (let i = 0; i < fields.results.length; i++) {
                        let field = fields.results[i];

                        // Add the dropdown items
                        control.items.push({
                            data: field,
                            text: field.Title,
                            value: field.InternalName
                        });
                    }

                    // Resolve the promise
                    resolve(control);
                }, reject);
            } else {
                // Resolve the promise
                resolve(control);
            }
        });
    }

    // Method to load the list fields
    let loadFields = (list: Types.SP.List) => {
        // Clear the dropdown
        _ddl.setItems([]);

        // Ensure a list exists
        if (list) {
            // Load the fields
            list.Fields().execute(fields => {
                let items: Array<Components.IDropdownItem> = [];

                // Parse the fields
                for (let i = 0; i < fields.results.length; i++) {
                    let field = fields.results[i];

                    // Add the dropdown items
                    items.push({
                        data: field, // The data property will store any reference data
                        text: field.Title,
                        value: field.InternalName
                    });
                }

                // Update the dropdown
                _ddl.setItems(items);
            });
        }
    }

    // Return the edit form configuration
    return {
        // Define the custom configuration for this webpart
        onRenderForm: (wpInfo, list) => {
            let cfg: Array<Components.IFormControlProps> = [];

            // Dropdown form component to display the list fields
            cfg.push({
                name: "List Field(s):",
                description: "These fields will be displayed in the table.",
                items: [],
                type: Components.FormControlTypes.MultiDropdown,
                onControlRendering: (control) => {
                    // Initialize the fields
                    return initFields(control, list as Types.SP.List);
                },
                onControlRendered: (control) => {
                    // Save a reference to the control
                    _ddl = control.get() as Components.IDropdown;
                }
            } as Components.IFormControlPropsDropdown);

            // Return the configuration
            return cfg;
        },
        // Reload the dropdown when a list is selected
        onListChanged: (wpInfo, list) => {
            // Load the fields
            loadFields(list as Types.SP.List);
        },
        // Update the configuration when saving
        onSave: (wpCfg: IBSDemoCfg) => {
            // Clear the field names
            wpCfg.FieldDisplayNames = [];
            wpCfg.FieldInternalNames = [];

            // Get the dropdown values
            let items = _ddl.getValue() as Array<Components.IDropdownItem>;
            for (let i = 0; i < items.length; i++) {
                let item = items[i];

                // The field can be referenced from the data
                let fld: Types.SP.IField = item.data;

                // Add the field name
                wpCfg.FieldDisplayNames.push(item.text);
                wpCfg.FieldInternalNames.push(item.value);
            }

            // Return the configuration
            return wpCfg;
        }
    }
})();
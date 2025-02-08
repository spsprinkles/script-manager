import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { MapperV2 } from "gd-sprest-bs/../gd-sprest/build/mapper";
import { fileExcel } from "gd-sprest-bs/build/icons/svgs/fileExcel";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { DataSource, IListItem } from "./ds";
import { ExportCSV } from "./exportCSV";
import { ProcessScript, IProcessResult } from "./process";

export class Forms {
    // Create form
    static create(onCreated: () => void) {
        // Show the create form
        DataSource.List.newForm({
            onCreateEditForm: props => {
                // Customize the form
                props.onControlRendering = this.customizeForm();
                return props;
            },
            onUpdate: () => {
                // Refresh the data
                DataSource.refresh().then(() => {
                    // Call the event
                    onCreated();
                });
            }
        });
    }

    // The methods field popover
    private static _items: Components.IDropdownItem[] = null;
    private static _selectedType: string = null;
    private static setItems(item: Components.IDropdownItem) {
        // Set the items, based on the value
        switch (item?.text) {
            case "File":
                this._selectedType = "driveItem";
                break;
            case "File Item":
            case "Item":
                this._selectedType = "listItem";
                break;
            case "List":
                this._selectedType = "list";
                break;
            case "Site":
                this._selectedType = "site";
                break;
            default:
                this._selectedType = "";
                break;
        }

        // Clear the items
        this._items = [{ text: "" }];

        // Get the library methods
        let lib = MapperV2[this._selectedType];
        if (lib) {
            // Parse the methods
            for (let methodName in lib) {
                let methodInfo = lib[methodName];

                // See if this is a method
                if (methodName != "properties") {
                    // See if arg names exist
                    this._items.push({
                        data: methodInfo,
                        text: methodName
                    });
                }
            }

            // See if this is the list type
            if (this._selectedType == "list") {
                // Add a clear method
                this._items.push({ text: "clear" });
            }

            // Sort the items
            this._items = this._items.sort((a, b) => {
                if (a.text > b.text) { return 1; }
                if (a.text < b.text) { return -1; }
                return 0;
            });
        }

        // Update the popover
        this._popover.setBody(Components.Dropdown({
            menuOnly: true,
            items: this._items,
            onChange: (item: Components.IDropdownItem) => {
                // Ensure an item was selected
                if (item) {
                    // Set the value
                    DataSource.List.EditForm.getControl("Method").setValue(item.text);
                }

                // See if there are any arguments
                let ctrlParms = DataSource.List.EditForm.getControl("Parameters");
                if (item.data?.argNames) {
                    // Set the params
                    ctrlParms.setValue(`[${item.data.argNames.join(', ')}]`);
                } else {
                    // Clear the params
                    ctrlParms.setValue("");
                }

                // Hide the popover
                this._popover.hide();
            }
        }).el);
    }

    // Customizes the form
    private static _popover: Components.IPopover = null;
    private static customizeForm(): (control: Components.IFormControlProps, field: Types.SP.Field) => void {
        return (ctrl, fld) => {
            // See if this is the method field
            if (fld.InternalName == "Method") {
                // Add a change event
                (ctrl as Components.IFormControlPropsDropdown).onChange = (item) => {
                    // Ensure an item is selected
                    if (item == null) { return; }

                    // Get the method params
                    this.getMethodParams(item.text);
                }

                // Add a rendered event
                ctrl.onControlRendered = (ctrl) => {
                    // Add a popover
                    this._popover = Components.Popover({
                        classNameBody: "graph-methods",
                        target: ctrl.textbox.elTextbox,
                        placement: Components.PopoverPlacements.BottomStart,
                        options: {
                            trigger: "click"
                        }
                    });

                    // Set the items if a value exists
                    let item: IListItem = DataSource.List.EditForm.getItem();
                    if (item?.ScriptType) {
                        // Set the items
                        this.setItems({ text: item.ScriptType });
                    }
                }
            }

            // See if it's the parameters field
            if (fld.InternalName == "Parameters") {
                // Add validation
                ctrl.onValidate = (ctrl, results) => {
                    // See if there is a value
                    if (ctrl.value) {
                        // Ensure it's in JSON format
                        try { JSON.parse(results.value); }
                        catch {
                            results.isValid = false;
                            results.invalidMessage = "The parameters need to be in JSON format.";
                        }
                    }

                    return results;
                }
            }

            // See if this is the script type
            if (fld.InternalName == "ScriptType") {
                // Clear the value
                ctrl.value = "";

                // Add a change event
                (ctrl as Components.IFormControlPropsDropdown).onChange = (item) => {
                    // Set the items
                    this.setItems(item);
                }
            }
        };
    }

    // Edit form
    static edit(itemId: number, onUpdated: () => void) {
        // Show the create form
        DataSource.List.editForm({
            itemId,
            onCreateEditForm: props => {
                // Customize the form
                props.onControlRendering = this.customizeForm();
                return props;
            },
            onUpdate: () => {
                // Refresh the data
                DataSource.refresh().then(() => {
                    // Call the event
                    onUpdated();
                });
            }
        });
    }

    // Get the method params
    private static getMethodParams(methodName: string) {
        // Get the lib
    }

    // Renders the summary dialog
    private static _table: DataTable = null;
    private static _results: IProcessResult[] = null;
    private static renderSummary(item: IListItem) {
        // Clear the modal
        Modal.clear();

        // Set the type
        Modal.setType(Components.ModalTypes.Full);

        // Prevent auto close
        Modal.setAutoClose(false);

        // Show the modal dialog
        Modal.setHeader("Processing Requests");

        // Render the table
        this._table = new DataTable({
            el: Modal.BodyElement,
            rows: [],
            columns: [
                {
                    name: "Error",
                    title: "Request Errored?"
                },
                {
                    name: "Message",
                    title: "Message"
                },
                {
                    name: "Output",
                    title: "Output",
                    onRenderCell: (el, col, item: IProcessResult) => {
                        // Try to convert the value
                        try {
                            // Convert the value
                            let value = JSON.parse(item.Output);

                            // Update the html
                            el.innerText = JSON.stringify(value, null, 2);
                        } catch { }
                    }
                }
            ]
        });

        // Set the footer
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Export to a CSV file",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconType: fileExcel(24, 24, "ExcelDocument", "mx-1"),
                        text: "Export",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Export the CSV
                            new ExportCSV(item.Title + "_results.csv", ["Error", "Message", "Output"], this._items);
                        }
                    }
                },
                {
                    content: "Close Window",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconClassName: "mx-1",
                        iconType: xSquare,
                        iconSize: 24,
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        }).el);

        // Show the modal
        Modal.show();
    }

    // Shows the upload form
    static upload(item: IListItem) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Upload CSV");

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [{
                name: "CSVFile",
                label: "CSV File",
                required: true,
                type: Components.FormControlTypes.File,
                errorMessage: "A file is required.",
                onValidate: (ctrl, results) => {
                    // See if a value exists
                    if (results.value) {
                        // Ensure it's a CSV file
                        results.isValid = /.csv$/i.test(results.value.name);
                        results.invalidMessage = "The file type must be a csv.";
                    }

                    // Return the results
                    return results;
                }
            }]
        });

        // Set the footer
        Components.Button({
            el: Modal.FooterElement,
            text: "Upload",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Ensure the form is valid
                if (form.isValid()) {
                    // Read the file in chunks to handle larger files
                    // Otherwise, you will get a memory error
                    let fileInfo = form.getValues()["CSVFile"];
                    let fileData = new Uint8Array(fileInfo.data);
                    let chunkSize = 10000;
                    let csv = "";
                    for (let i = 0; i < fileData.length; i += chunkSize) {
                        csv += String.fromCharCode.apply(null, fileData.slice(i, i + chunkSize));
                    }

                    // Clear the results
                    this._results = [];

                    // Render the summary table
                    this.renderSummary(item);

                    // Process the rows
                    let rows = csv.split('\n');
                    new ProcessScript(item, rows, (result) => {
                        // Append the result
                        this._results.push(result);

                        // Refresh the table
                        this._table.refresh(this._results);

                        // Update the header
                        if (rows.length == this._results.length) {
                            Modal.setHeader(`${item.Title} Results: ${rows.length} Completed`);
                        } else {
                            Modal.setHeader(`Processed (${this._results.length} of ${rows.length}) Requests`);
                        }
                    });
                }
            }
        });

        // Show the modal
        Modal.show();
    }
}
import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, Types, Web } from "gd-sprest-bs";
import { fileExcel } from "gd-sprest-bs/build/icons/svgs/fileExcel";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { DataSource, IListItem } from "./ds";
import { ExportCSV } from "./exportCSV";
import { ProcessScript, IProcessResult } from "./process";
import Strings from "./strings";
import { Templates } from "./templates";

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
    private static setItems(item: Components.IDropdownItem) {
        // Set the items, based on the value
        switch (item?.text) {
            case "File":
                this._items = Templates.FileMethods;
                break;
            case "Item":
                this._items = Templates.ItemMethods;
                break;
            case "List":
                this._items = Templates.ListMethods;
                break;
            case "Site":
                this._items = Templates.SiteMethods;
                break;
            default:
                this._items = [];
                break;
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
                // Add a rendered event
                ctrl.onControlRendered = (ctrl) => {
                    // Add a popover
                    this._popover = Components.Popover({
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

    // Process form
    static process(item: IListItem): PromiseLike<IProcessResult[]> {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading CSV");
        LoadingDialog.setBody("Loading the csv for this item.");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the attachment files
            item.AttachmentFiles().execute(files => {
                // Ensure the file exists
                if (files.results[0]) {
                    // Read the file
                    this.readFile(document.location.origin + files.results[0].ServerRelativeUrl).then(csv => {
                        // Hide the dialog
                        LoadingDialog.hide();

                        // Process the rows
                        new ProcessScript(item, csv.split('\n'), (results) => {
                            // Show the results
                            this.renderSummary(item, results);

                            // Hide the loading dialog
                            LoadingDialog.hide();
                        });
                    }, () => {
                        // Clear the modal
                        Modal.clear();

                        // Set the header
                        Modal.setHeader("Load Error");

                        // Set the body
                        Modal.setBody("There was an error loading the csv file.");

                        // Show the modal
                        Modal.show();

                        // Hide the loading dialog and reject the request
                        LoadingDialog.hide();
                        reject();
                    });
                } else {
                    // Clear the modal
                    Modal.clear();

                    // Set the header
                    Modal.setHeader("Load Error");

                    // Set the body
                    Modal.setBody("There is no csv file associated with this item. Please upload a csv file first.");

                    // Show the modal
                    Modal.show();

                    // Hide the loading dialog and reject the request
                    LoadingDialog.hide();
                    reject();
                }

            });
        });
    }

    // Reads the file attachment and returns a string
    private static readFile(url: string): PromiseLike<string> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Read the file
            Web(Strings.SourceUrl).getFileByUrl(url).execute(file => {
                // Read the content
                file.content().execute(data => {
                    // Resolve the request
                    resolve(String.fromCharCode.apply(null, new Uint8Array(data)));
                });
            }, () => {
                // Clear the modal
                Modal.clear();

                // Set the header
                Modal.setHeader("Load Error");

                // Set the body
                Modal.setBody("There was an error loading the csv file.");

                // Show the modal
                Modal.show();

                // Reject the request
                reject();
            });
        });
    }

    // Renders the summary dialog
    private static renderSummary(item: IListItem, results: IProcessResult[]) {
        // Clear the modal
        Modal.clear();

        // Set the type
        Modal.setType(Components.ModalTypes.Full);

        // Prevent auto close
        Modal.setAutoClose(false);

        // Show the modal dialog
        Modal.setHeader(item.Title + " Results");

        // Render the table
        new DataTable({
            el: Modal.BodyElement,
            rows: results,
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
                            new ExportCSV(item.Title + "_results.csv", ["Error", "Message", "Output"], results);
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

    // Update form
    static update(item: IListItem) {
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
                    // Show a loading dialog
                    LoadingDialog.setHeader("Uploading CSV");
                    LoadingDialog.setBody("Removing the previous file...");
                    LoadingDialog.show();

                    // Load the current attachments
                    item.AttachmentFiles().execute(files => {
                        // Parse the current files
                        Helper.Executor(files.results, file => {
                            // Delete the file
                            return file.delete().execute();
                        }).then(() => {
                            // Update the dialog
                            LoadingDialog.setBody("Uploading the new file...");

                            // Upload the file
                            let fileInfo = form.getValues()["CSVFile"];
                            item.AttachmentFiles().add(fileInfo.name, fileInfo.data).execute(
                                () => {
                                    // Refresh the item
                                    DataSource.refresh(item.Id).then(() => {
                                        // Close the loading dialogs
                                        LoadingDialog.hide();
                                        Modal.hide();
                                    });
                                },
                                () => {
                                    // Update the body
                                    Modal.clear();
                                    Modal.setHeader("File Upload Failed");
                                    Modal.setBody("The file failed to upload. Please refresh the page and try again.");

                                    // Hide the dialog
                                    LoadingDialog.hide();
                                }
                            );
                        });
                    });
                }
            }
        });

        // Show the modal
        Modal.show();
    }
}
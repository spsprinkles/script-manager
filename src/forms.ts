import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, Web } from "gd-sprest-bs";
import { fileExcel } from "gd-sprest-bs/build/icons/svgs/fileExcel";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { DataSource, IListItem } from "./ds";
import { ExportCSV } from "./exportCSV";
import { ProcessScript, IProcessResult } from "./process";
import Strings from "./strings";

export class Forms {
    // Create form
    static create(onCreated: () => void) {
        // Show the create form
        DataSource.List.newForm({
            onGetListInfo: props => {
                // Load attachments
                props.loadAttachments = true;
                return props;
            },
            onCreateEditForm: props => {
                // Show the attachments
                props.displayAttachments = true;
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

    // Edit form
    static edit(itemId: number, onUpdated: () => void) {
        // Show the create form
        DataSource.List.editForm({
            itemId,
            onGetListInfo: props => {
                // Load attachments
                props.loadAttachments = true;
                return props;
            },
            onCreateEditForm: props => {
                // Show the attachments
                props.displayAttachments = true;
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
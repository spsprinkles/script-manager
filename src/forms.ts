import { LoadingDialog, Modal } from "dattatable";
import { Types, Web } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";
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
            // Ensure the file exists
            if (item.AttachmentFiles[0]) {
                // Read the file
                this.readFile(item.AttachmentFiles[0].ServerRelativeUrl).then(csv => {
                    // Process the rows
                    new ProcessScript(item, csv.split('\n'), (results) => {
                        // Show the results
                        // TODO

                        // Hide the dialod
                        LoadingDialog.hide();
                    });
                }, () => {
                    // Hide the loading dialog and reject the request
                    LoadingDialog.hide();
                    reject();
                });
            }
        });
    }

    // Reads the file attachment and returns a string
    private static readFile(url: string): PromiseLike<string> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Read the file
            Web(Strings.SourceUrl).getFileByUrl(url).content().execute(data => {
                // Resolve the request
                resolve(String.fromCharCode.apply(null, new Uint8Array(data)));
            }, () => {
                // Clear the modal
                Modal.clear();

                // Set the header
                Modal.setHeader("Load Error");

                // Set the body
                Modal.setBody("There was an error loading the csv file.");

                // Show the modal
                Modal.show();
            });
        });
    }

    // Update form
    static update(item: IListItem) {
        // TODO
    }
}
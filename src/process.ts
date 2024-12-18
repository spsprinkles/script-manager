import { LoadingDialog } from "dattatable";
import { Helper, v2 } from "gd-sprest-bs";
import { IListItem } from "./ds";
import { Templates } from "./templates";

export interface IProcessResult {
    Error: boolean;
    Message?: string;
    Output?: string;
}

/**
 * Process Script
 */
export class ProcessScript {
    private _item: IListItem = null;
    private _onCompleted: (results: IProcessResult[]) => void;
    private _rows: string[] = null;

    constructor(item: IListItem, rows: string[], onCompleted: (results: IProcessResult[]) => void) {
        this._item = item;
        this._onCompleted = onCompleted;
        this._rows = rows;

        // Remove the first row
        this._rows.splice(0, 1);

        // Process the rows
        this.process();
    }

    // Method to execute the graph request
    private execute(obj, method: string, params: string): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Try to convert the parameters
            let args = null;
            try {
                args = JSON.parse(params);
                args = typeof (args) === "object" && typeof (args.length) === "number" ? args : [args];
            }
            catch { args = []; }

            // Ensure the method exists
            if (method) {
                // Execute the method
                obj[method](...args).execute(resp => {
                    // Resolve the request
                    resolve({
                        Error: false,
                        Message: "Method was executed successfully.",
                        Output: resp.response
                    });
                }, err => {
                    // Resolve the request
                    resolve({
                        Error: true,
                        Message: "Error executing the method.",
                        Output: err.response
                    });
                });
            } else {
                // Resolve the request
                resolve({
                    Error: true,
                    Message: "Unable to determine the method."
                });
            }
        });
    }

    // Processes the rows
    private process() {
        let results: IProcessResult[] = [];

        // Show a loading dialog
        LoadingDialog.setHeader("Processing " + this._item.Title + " Script");
        LoadingDialog.setBody("Running the script...");
        LoadingDialog.show();

        // Process the rows
        let counter = 0;
        Helper.Executor(this._rows, row => {
            // Return a promise
            return new Promise(resolve => {
                // Update the loading dialog
                LoadingDialog.setBody(`Processing ${++counter} of ${this._rows.length}`);

                // Get the row data
                let data = row.split(',');
                for (let i = 0; i < data.length; i++) { data[i] = data[i].trim(); }

                // Process based on the type
                switch (this._item.ScriptType) {
                    case "File":
                        this.processFile(data).then(result => { results.push(result); resolve(null); });
                        break;
                    case "List":
                        this.processList(data).then(result => { results.push(result); resolve(null); });
                        break;
                    case "Site":
                        this.processSite(data).then(result => { results.push(result); resolve(null); });
                        break;
                }
            });
        }).then(() => {
            // Call the event
            this._onCompleted(results);
        });
    }

    // Process the file
    private processFile(row: string[]): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Get the file
            v2.sites.getFile({
                fileUrl: row[Templates.FileColumns.FileUrl],
                listName: row[Templates.FileColumns.ListName],
                siteUrl: row[Templates.FileColumns.SiteUrl]
            }).then(file => {
                // Execute the method
                this.execute(file, row[Templates.FileColumns.Method] || this._item.Method, row[Templates.FileColumns.Parameters] || this._item.Parameters).then(resolve);
            }, err => {
                // Resolve the request
                resolve({
                    Error: true,
                    Message: "Error getting the file.",
                    Output: err.response || err
                });
            });
        });
    }

    // Process the list
    private processList(row: string[]): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Get the list
            v2.sites.getList({
                listId: row[Templates.ListColumns.ListId],
                listName: row[Templates.ListColumns.ListName],
                siteUrl: row[Templates.ListColumns.SiteUrl]
            }).then(list => {
                // Execute the method
                this.execute(list, row[Templates.ListColumns.Method] || this._item.Method, row[Templates.ListColumns.Parameters] || this._item.Parameters).then(resolve);
            }, err => {
                // Resolve the request
                resolve({
                    Error: true,
                    Message: "Error getting the list.",
                    Output: err.response || err
                });
            });
        });
    }

    // Process the site
    private processSite(row: string[]): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Get the list
            var site = v2.sites.getSite(row[Templates.SiteColumns.SiteUrl]);

            // Execute the method
            this.execute(site, row[Templates.SiteColumns.Method] || this._item.Method, row[Templates.SiteColumns.Parameters] || this._item.Parameters).then(resolve);
        });
    }
}
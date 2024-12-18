import { LoadingDialog } from "dattatable";
import { Helper, v2 } from "gd-sprest-bs";
import { IListItem } from "./ds";
import { Templates } from "./templates";

export interface IProcessResult {
    Error: boolean;
    Message?: string;
    Output: string;
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

        // Process the rows
        this.process();
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
            // Update the loading dialog
            LoadingDialog.setBody(`Processing ${++counter} of ${this._rows.length}`);

            // Get the row data
            let data = row.split(',');
            for (let i = 0; i < data.length; i++) { data[i] = data[i].trim(); }

            // Process based on the type
            switch (this._item.ScriptType) {
                case "File":
                    this.processFile(data).then(result => { results.push(result); });
                    break;
                case "List":
                    this.processList(data).then(result => { results.push(result); });
                    break;
                case "Site":
                    this.processSite(data).then(result => { results.push(result); });
                    break;
            }
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
                file[row[Templates.FileColumns.Method || this._item.Method]].apply(file, row[Templates.FileColumns.Parameters] || this._item.Parameters).execute(resp => {
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
            }, err => {
                // Resolve the request
                resolve({
                    Error: true,
                    Message: "Error getting the file.",
                    Output: err.response
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
                list[row[Templates.ListColumns.Method || this._item.Method]].apply(list, row[Templates.ListColumns.Parameters] || this._item.Parameters).execute(resp => {
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
            }, err => {
                // Resolve the request
                resolve({
                    Error: true,
                    Message: "Error getting the list.",
                    Output: err.response
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
            site[row[Templates.SiteColumns.Method || this._item.Method]].apply(site, row[Templates.SiteColumns.Parameters] || this._item.Parameters).execute(resp => {
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
        });
    }
}
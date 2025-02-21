import { Helper, v2, Web } from "gd-sprest-bs";
import { IListItem } from "./ds";
import { Templates } from "./templates";

export interface IProcessResult {
    SourceRow: object;
    Error: boolean;
    Message?: string;
    Output?: string;
}

/**
 * Process Script
 */
export class ProcessScript {
    private _item: IListItem = null;
    private _onProcessedItem: (results: IProcessResult) => void;
    private _rows: string[] = null;

    constructor(item: IListItem, rows: string[], onProcessedItem: (result: IProcessResult) => void) {
        this._item = item;
        this._onProcessedItem = onProcessedItem;
        this._rows = rows;

        // Remove the first row
        this._rows.splice(0, 1);

        // See if the last row is empty
        if (this._rows[this._rows.length - 1] == "") {
            // Remove the last row
            this._rows.splice(this._rows.length - 1, 1);
        }

        // Process the rows
        this.process();
    }

    // Method to execute the graph request
    private execute(srcRow, obj, method: string, params: string): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Try to convert the parameters
            let args = null;
            try {
                args = JSON.parse(params);
                args = typeof (args) === "object" && typeof (args.length) === "number" ? args : [args];
            }
            catch { args = []; }

            // See if the method exists
            if (method && obj[method] == null) {
                // Resolve the request
                resolve({
                    SourceRow: srcRow,
                    Error: true,
                    Message: "Error: The method does not exist for this object type.",
                    Output: "Unable to process request."
                });
            } else {
                // Execute the method
                (method ? obj[method](...args) : obj).execute(resp => {
                    // Resolve the request
                    resolve({
                        SourceRow: srcRow,
                        Error: false,
                        Message: "Method was executed successfully.",
                        Output: resp?.response
                    });
                }, err => {
                    // Resolve the request
                    resolve({
                        SourceRow: srcRow,
                        Error: true,
                        Message: "Error executing the method.",
                        Output: err?.response
                    });
                });
            }
        });
    }

    // Processes the rows
    private process() {
        // Process the rows
        Helper.Executor(this._rows, row => {
            // Return a promise
            return new Promise(resolve => {
                // Set the complete method
                let onComplete = (item => {
                    // Call the event
                    this._onProcessedItem(item);

                    // Check the next item
                    resolve(item);
                });

                // Get the row data
                let data = [];
                let currentValue = "";
                let insideQuote = false;
                for (let i = 0; i < row.length; i++) {
                    let appendFl = false;

                    if (row[i] === '"') {
                        // See if this is the start
                        if (!insideQuote) {
                            // Set the flag and continue
                            insideQuote = true;
                            continue;
                        } else {
                            // Set the flag
                            appendFl = true;
                        }
                    } else if (row[i] === ',') {
                        // See if we are not inside a quote
                        if (!insideQuote) {
                            // Set the flag
                            appendFl = true;
                        }
                    }

                    // See if we are appending the value
                    if (appendFl) {
                        // Append the current value, reset the value/flag and continue
                        data.push(currentValue.trim());
                        currentValue = "";
                        insideQuote = false;
                    } else {
                        // Append the value
                        currentValue += row[i];
                    }
                }

                // Process based on the type
                switch (this._item.ScriptType) {
                    case "File":
                        this.processFile(data, false).then(onComplete);
                        break;
                    case "File Item":
                        this.processFile(data, true).then(onComplete);
                        break;
                    case "Item":
                        this.processItem(data).then(onComplete);
                        break;
                    case "List":
                        this.processList(data).then(onComplete);
                        break;
                    case "Site":
                        this.processSite(data).then(onComplete);
                        break;
                }
            });
        });
    }

    // Process the file
    private processFile(row: string[], itemFl: boolean): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Get the file
            v2.sites.getFile({
                fileUrl: row[Templates.FileColumns.FileUrl],
                listName: row[Templates.FileColumns.ListName],
                siteUrl: row[Templates.FileColumns.SiteUrl]
            }).then(file => {
                // See if we are applying it to the item
                if (itemFl) {
                    // Execute the method
                    this.execute(row, file.listItem(), row[Templates.FileColumns.Method] || this._item.Method, row[Templates.FileColumns.Parameters] || this._item.Parameters).then(resolve);
                } else {
                    // Execute the method
                    this.execute(row, file, row[Templates.FileColumns.Method] || this._item.Method, row[Templates.FileColumns.Parameters] || this._item.Parameters).then(resolve);
                }
            }, err => {
                // Resolve the request
                resolve({
                    SourceRow: row,
                    Error: true,
                    Message: "Error getting the file.",
                    Output: err?.response || err
                });
            });
        });
    }

    // Process the item
    private processItem(row: string[]): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Get the file
            v2.sites.getList({
                listName: row[Templates.ItemColumns.ListName],
                siteUrl: row[Templates.ItemColumns.SiteUrl]
            }).then(list => {
                let item = list.items(row[Templates.ItemColumns.ItemID]);

                // Try to execute the method
                try {
                    this.execute(row, item, row[Templates.ItemColumns.Method] || this._item.Method, row[Templates.ItemColumns.Parameters] || this._item.Parameters).then(resolve);
                } catch {
                    // Resolve the request
                    resolve({
                        SourceRow: row,
                        Error: true,
                        Message: "Error executing the method.",
                        Output: "The parameters are not in the correct format for this method."
                    });
                }
            }, err => {
                // Resolve the request
                resolve({
                    SourceRow: row,
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
            // See if this is the clear method
            let method = row[Templates.ListColumns.Method] || this._item.Method;
            if (method == "clear") {
                var list = Web(row[Templates.ListColumns.SiteUrl]).Lists(row[Templates.ListColumns.ListName]);

                // Get the items
                list.Items().query({
                    GetAllItems: true,
                    Select: ["id"],
                    Top: 5000
                }).execute(items => {
                    // Delete the items
                    for (let i = 0; i < items.results.length; i++) {
                        // Delete the item
                        list.Items(items.results[i].Id).delete().batch();
                    }

                    // Execute the batch job
                    list.execute(resp => {
                        // Resolve the request
                        resolve({
                            SourceRow: row,
                            Error: false,
                            Message: "The list has been cleared.",
                            Output: resp["response"]
                        });
                    });
                });
            } else {
                // Get the list
                v2.sites.getList({
                    listId: row[Templates.ListColumns.ListId],
                    listName: row[Templates.ListColumns.ListName],
                    siteUrl: row[Templates.ListColumns.SiteUrl]
                }).then(list => {
                    /* Comment out until we confirm that batch is available in v2 REST
                    let method = row[Templates.ListColumns.Method] || this._item.Method;

                    // See if this is the clear method
                    if (method == "clear") {
                        // Get the items
                        list.items().query({
                            GetAllItems: true,
                            Select: ["id"],
                            Top: 5000
                        }).execute(items => {
                            // Delete the items
                            for (let i = 0; i < items.results.length; i++) {
                                // Delete the item
                                list.Items(items.results[i].id).delete().batch();
                            }

                            // Execute the batch job
                            list.execute(resp => {
                                // Resolve the request
                                resolve({
                                    Error: false,
                                    Message: "The list has been cleared.",
                                    Output: resp["response"]
                                });
                            });
                        });
                    } else {
                        // Execute the method
                        this.execute(list, method, row[Templates.ListColumns.Parameters] || this._item.Parameters).then(resolve);
                    }
                    */
                    // Execute the method
                    this.execute(row, list, method, row[Templates.ListColumns.Parameters] || this._item.Parameters).then(resolve);
                }, err => {
                    // Resolve the request
                    resolve({
                        SourceRow: row,
                        Error: true,
                        Message: "Error getting the list.",
                        Output: err.response || err
                    });
                });
            }
        });
    }

    // Process the site
    private processSite(row: string[]): PromiseLike<IProcessResult> {
        // Return a promise
        return new Promise(resolve => {
            // Get the list
            v2.sites.getSite(row[Templates.SiteColumns.SiteUrl]).then(site => {
                // Execute the method
                this.execute(row, site, row[Templates.SiteColumns.Method] || this._item.Method, row[Templates.SiteColumns.Parameters] || this._item.Parameters).then(resolve);
            });
        });
    }
}
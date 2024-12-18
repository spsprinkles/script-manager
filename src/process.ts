import { LoadingDialog } from "dattatable";
import { Helper } from "gd-sprest-bs";
import { IListItem } from "./ds";

export interface IProcessResult {
    // TODO
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
        // Show a loading dialog
        LoadingDialog.setHeader("Processing " + this._item.Title + " Script");
        LoadingDialog.setBody(`Processing 1 of ${this._rows.length}`);
        LoadingDialog.show();

        // Process the rows
        Helper.Executor(this._rows, row => {
            // Process based on the type
            switch (this._item.ScriptType) {
                case "File":
                    break;
                case "List":
                    break;
                case "Site":
                    break;
            }
        });
    }
}
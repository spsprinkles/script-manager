export class Templates {
    // File
    private static File = `"Site Url", "File Url", "Method", "Parameters"`;

    // List
    private static List = `"Site Url", "List Name", "Method", "Parameters"`;

    // Site
    private static Site = `"Site Url", "Method", "Parameters"`;

    // Downloads the the template csv
    static download(title: string, templateType: "file" | "list" | "site") {
        // Set the file name and template
        let filename = "";
        let template = "";
        switch (templateType) {
            case "file":
                filename = title + "_file.csv";
                template = this.File;
                break;
            case "list":
                filename = title + "_list.csv";
                template = this.List;
                break;
            case "site":
                filename = title + "_site.csv";
                template = this.Site;
                break;
            // Invalid
            default:
                return;
        }

        // See if this is IE or Mozilla
        if (Blob && navigator && navigator["msSaveBlob"]) {
            // Download the file
            navigator["msSaveBlob"](new Blob([template], { type: "data:text/csv;charset=utf-8;" }), filename);
        } else {
            // Generate an anchor
            var anchor = document.createElement("a");
            anchor.download = filename;
            anchor.href = "data:text/csv;charset=utf-8," + encodeURIComponent(template);
            anchor.target = "__blank";

            // Download the file
            anchor.click();
        }
    }
}
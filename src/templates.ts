export class Templates {
    // File
    private static File = `Site Url,List Name,File Url,Method,Parameters`;
    static FileColumns = {
        SiteUrl: 0,
        ListName: 1,
        FileUrl: 2,
        Method: 3,
        Parameters: 4
    }

    // List
    private static List = `Site Url,List ID,List Name,Method,Parameters`;
    static ListColumns = {
        SiteUrl: 0,
        ListId: 1,
        ListName: 2,
        Method: 3,
        Parameters: 4
    }

    // Site
    private static Site = `Site Url,Method,Parameters`;
    static SiteColumns = {
        SiteUrl: 0,
        Method: 1,
        Parameters: 2
    }

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
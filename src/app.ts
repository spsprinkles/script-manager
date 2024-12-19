import { Dashboard } from "dattatable";
import { Components } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";
import { Forms } from "./forms";
import { InstallDialog } from "./install";
import Strings from "./strings";
import { Templates } from "./templates";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        let dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            navigation: {
                title: Strings.ProjectName,
                showFilter: false,
                items: [
                    {
                        className: "btn-outline-light",
                        text: "Create",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            Forms.create(() => {
                                // Refresh the table
                                dashboard.refresh(DataSource.ListItems);
                            });
                        }
                    }
                ],
                itemsEnd: [
                    {
                        className: "btn-outline-light me-2",
                        text: "Templates",
                        isButton: true,
                        items: [
                            {
                                text: "File",
                                onClick: () => { Templates.download("template", "file"); }
                            },
                            {
                                text: "Item",
                                onClick: () => { Templates.download("template", "item"); }
                            },
                            {
                                text: "List",
                                onClick: () => { Templates.download("template", "list"); }
                            },
                            {
                                text: "Site",
                                onClick: () => { Templates.download("template", "site"); }
                            }
                        ]
                    },
                    {
                        className: "btn-outline-light me-2",
                        text: "Settings",
                        isButton: true,
                        onClick: () => {
                            // Show the install modal
                            InstallDialog.show(true);
                        }
                    }
                ]
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.ListItems,
                // Update the default datatables.net properties
                onRendering: dtProps => {
                    // Remove the ability to sort and search on the 1st column
                    dtProps.columnDefs = [
                        {
                            "targets": 4,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 1st column by default; ascending
                    dtProps.order = [[1, "asc"]];

                    // Return the datatables.net properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "Title",
                        title: "Title"
                    },
                    {
                        name: "ScriptType",
                        title: "Script Type"
                    },
                    {
                        name: "Method",
                        title: "Method"
                    },
                    {
                        name: "Parameters",
                        title: "Parameters"
                    },
                    {
                        name: "Actions",
                        title: "",
                        onRenderCell: (el, col, item: IListItem) => {
                            // Render the actions
                            Components.TooltipGroup({
                                el,
                                isSmall: true,
                                tooltips: [
                                    {
                                        content: "Modifies the script method and properties.",
                                        btnProps: {
                                            text: "Modify",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the edit form
                                                Forms.edit(item.Id, () => {
                                                    // Refresh the table
                                                    dashboard.refresh(DataSource.ListItems);
                                                });
                                            }
                                        }
                                    },
                                    {
                                        content: "Uploads a new csv file to process.",
                                        btnProps: {
                                            text: "Upload CSV",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the update form
                                                Forms.update(item);
                                            }
                                        }
                                    },
                                    {
                                        content: "Processes the csv file.",
                                        btnProps: {
                                            text: "Process",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the process form
                                                Forms.process(item);
                                            }
                                        }
                                    }
                                ]
                            })
                        }
                    }
                ]
            }
        });
    }
}
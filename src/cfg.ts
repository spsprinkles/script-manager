import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            ContentTypes: [{
                Name: "Item",
                FieldRefs: [
                    "Title",
                    "ScriptType",
                    "Method",
                    "Parameters"
                ]
            }],
            CustomFields: [
                {
                    name: "Method",
                    title: "Method",
                    type: Helper.SPCfgFieldType.Text,
                    required: true
                },
                {
                    name: "Parameters",
                    title: "Parameters",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScriptType",
                    title: "Script Type",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "File",
                    required: true,
                    choices: [
                        "File", "File Item", "Item", "List", "Site"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "ScriptType", "Method", "Parameters"
                    ]
                }
            ]
        }
    ]
});
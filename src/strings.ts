import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "script-manager",
    GlobalVariable: "ScriptManager",
    Lists: {
        Main: "Scripts"
    },
    ProjectName: "Script Manager",
    ProjectDescription: "Runs files/items against the REST API.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.6"
};
export default Strings;
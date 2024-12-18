import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

// Reference the solution
import "main-lib";
declare const ScriptManager: {
  render: (el: HTMLElement, context: WebPartContext) => void;
  updateTheme: (themeInfo: IReadonlyTheme) => void;
};

export interface IScriptManagerWebPartProps {
  description: string;
}

export default class ScriptManagerWebPart extends BaseClientSideWebPart<IScriptManagerWebPartProps> {

  public render(): void {
    // Clear the element
    while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

    // Render the application
    ScriptManager.render(this.domElement, this.context);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    ScriptManager.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}

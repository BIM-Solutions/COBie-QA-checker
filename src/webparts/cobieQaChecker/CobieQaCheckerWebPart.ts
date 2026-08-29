import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CobieQaCheckerWebPartStrings';
import { App } from './app/App';
import { SpFileService, SpRunHistoryService } from './services/sharepoint';
import type { CheckerServices } from './services/CheckerServices';

export interface ICobieQaCheckerWebPartProps {
  /** Library the file picker reads from. Blank means "Shared Documents". */
  sourceLibraryUrl: string;
  /** Library exported reports are written to. Blank means the source library. */
  reportLibraryUrl: string;
  /** Whether each run is recorded to the site's history list. */
  recordHistory: boolean;
}

/**
 * Web part entry point.
 *
 * Responsibilities are deliberately small: build the services from the page
 * context and mount `<App/>`. Everything else lives in app/, features/ and
 * services/, so the SPFx surface stays a thin edge that a test never has to
 * stand up.
 */
export default class CobieQaCheckerWebPart extends BaseClientSideWebPart<ICobieQaCheckerWebPartProps> {

  private _services: CheckerServices | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._buildServices();
  }

  /**
   * Rebuilds the services so a property-pane change - a different source
   * library, say - takes effect without the page being reloaded.
   */
  protected onPropertyPaneFieldChanged(): void {
    this._buildServices();
    this.render();
  }

  private _buildServices(): void {
    const web = this.context.pageContext.web;
    const files = new SpFileService(this.context.spHttpClient, web.absoluteUrl, web.serverRelativeUrl);

    const sourceFolder = files.resolveFolder(this.properties.sourceLibraryUrl);

    this._services = {
      files,
      history: new SpRunHistoryService(this.context.spHttpClient, web.absoluteUrl),
      sourceFolder,
      // Reports land beside the file they describe unless told otherwise. The
      // alternative - defaulting to a separate library that may not exist -
      // fails at the last step of the workflow, which is the worst place to fail.
      reportFolder: (this.properties.reportLibraryUrl || '').trim() === ''
        ? sourceFolder
        : files.resolveFolder(this.properties.reportLibraryUrl),
      currentUser: this.context.pageContext.user.email || this.context.pageContext.user.displayName,
      // Undefined on a web part added before this property existed, and history
      // on by default is the useful behaviour; `!== false` rather than a truthy
      // test keeps an explicit "off" honoured.
      historyEnabled: this.properties.recordHistory !== false
    };
  }

  public render(): void {
    if (!this._services) { return; }
    ReactDom.render(React.createElement(App, { services: this._services }), this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) { return; }
    const { semanticColors } = currentTheme;
    if (!semanticColors) { return; }
    // The web part inherits the site's theme rather than imposing its own; these
    // are the tokens the stylesheet reads.
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
    this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || null);
    this.domElement.style.setProperty('--link', semanticColors.link || null);
    this.domElement.style.setProperty('--dividerLine', semanticColors.bodyDivider || null);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.LibraryGroupName,
              groupFields: [
                PropertyPaneTextField('sourceLibraryUrl', {
                  label: strings.SourceLibraryFieldLabel,
                  description: strings.SourceLibraryFieldDescription,
                  placeholder: 'Shared Documents'
                }),
                PropertyPaneTextField('reportLibraryUrl', {
                  label: strings.ReportLibraryFieldLabel,
                  description: strings.ReportLibraryFieldDescription
                })
              ]
            },
            {
              groupName: strings.HistoryGroupName,
              groupFields: [
                PropertyPaneToggle('recordHistory', {
                  label: strings.RecordHistoryFieldLabel,
                  onText: strings.RecordHistoryOn,
                  offText: strings.RecordHistoryOff
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

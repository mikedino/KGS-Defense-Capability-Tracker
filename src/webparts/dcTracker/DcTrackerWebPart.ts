import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import App from './components/main';
import { ContextInfo } from 'gd-sprest-bs';
import { InstallationRequired } from 'dattatable';
import { Configuration } from './components/data/cfg';
import { Security } from './components/services/Security';

import Strings, { setContext } from './components/common/strings';
import * as strings from 'DcTrackerWebPartStrings';
import { formatError } from './components/common/utils';
import { IDCTrackerProps } from './components/common/props';

export interface IDcTrackerWebPartProps {
  description: string;
}

export default class DcTrackerWebPart extends BaseClientSideWebPart<IDcTrackerWebPartProps> {

  private _renderContent(element: React.ReactElement): void {
    ReactDom.render(element, this.domElement);
  }

  public _renderTracker(): void {
    const element: React.ReactElement<IDCTrackerProps> = React.createElement(
      App,
      {
        appDescription: this.properties.description,
        context: this.context,
      }
    );

    this._renderContent(element);
  }

  private _renderError(error: string): void {
    this._renderContent(React.createElement(
      'div',
      { className: 'pad', 'data-dct-startup-error': true },
      React.createElement('h3', null, `${Strings.ProjectName} Application`),
      React.createElement('p', null, 'The application could not finish starting. Please contact your administrator.'),
      React.createElement('p', null, `Error Message: ${error}`)
    ));
  }

  public async render(): Promise<void> {

    // set the context
    setContext(this.context);

    // set the config URL
    Configuration.setWebUrl(this.context.pageContext.web.serverRelativeUrl);

    try {
      console.log(`[${Strings.ProjectName}] Initialize Security Class`);
      await Security.init();
    } catch (err) {
      const message = formatError(err);
      console.error(`[${Strings.ProjectName}] Security initialization error:`, message, err);
      this._renderError(`Unable to initialize application security. ${message}`);
      return;
    }

    // DCT administrators manage application data, but only SharePoint site installers
    // should run the list/schema check. The installer requires web-level permissions.
    const canCheckInstallation = ContextInfo.isSiteAdmin || ContextInfo.isSiteOwner;

    if (canCheckInstallation) {
      try {
        console.log(`[${Strings.ProjectName}] Checking SharePoint configuration`);
        const requiresInstall = await InstallationRequired.requiresInstall({ cfg: Configuration });

        if (requiresInstall) {
          InstallationRequired.showDialog();
          return;
        }
      } catch (err) {
        const message = formatError(err);
        console.error(`[${Strings.ProjectName}] Configuration check error:`, message, err);
        this._renderError(`Unable to verify the SharePoint configuration. ${message}`);
        return;
      }
    }

    console.log(`[${Strings.ProjectName}] Installation complete. Render web part.`);
    this._renderTracker();

  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

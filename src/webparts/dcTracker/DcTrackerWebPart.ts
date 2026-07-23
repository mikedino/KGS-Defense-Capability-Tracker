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
import { InstallationModal } from './components/services/Installer';
import { IDCTrackerProps } from './components/common/props';

export interface IDcTrackerWebPartProps {
  description: string;
}

export default class DcTrackerWebPart extends BaseClientSideWebPart<IDcTrackerWebPartProps> {


  public _renderTracker(): void {
    const element: React.ReactElement<IDCTrackerProps> = React.createElement(
      App,
      {
        appDescription: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _renderError(error: string): void {
    this.domElement.innerHTML = `
      <div class="pad">
        <h3>${Strings.ProjectName} Application</h3>
        <p>This solution requires setup. Please contact your administrator.</p>
        <p>Error Message: ${error}</p>
      </div>`;
  }

  public async render(): Promise<void> {

    // set the context
    setContext(this.context);

    // set the config URL
    Configuration.setWebUrl(this.context.pageContext.web.serverRelativeUrl);

    try {

      console.log(`[${Strings.ProjectName}] Initialize Security Class`);
      await Security.init();

      const inAdminGroup = await Security.hasPermissions();

      console.log(`[${Strings.ProjectName}] Checking user permissions`);
      const hasFullControl = ContextInfo.isSiteAdmin || ContextInfo.isSiteOwner || inAdminGroup;

      if (hasFullControl) {
        //only Admins check if install is reqd
        const requiresInstall = await InstallationRequired.requiresInstall({ cfg: Configuration });
        if (requiresInstall) {
          InstallationRequired.showDialog();
        } else {
          console.log(`[${Strings.ProjectName}] Installation complete. Render web part.`);
          this._renderTracker();
          return;
        }
      } else {
        console.log(`[${Strings.ProjectName}] Installation complete. Render web part.`);
        this._renderTracker();
        return;
      }

    } catch (err) {
      
      console.error(`[${Strings.ProjectName}] Installation error:`, formatError(err));

      console.warn(`[${Strings.ProjectName}] Attempting to run the install...`);

      // If init fails (e.g., groups missing), run the installation modal
      // fallback - SHOULD BE A ONE TIME RUN ON FIRST INSTALL
      try {
        await InstallationModal.show(true, () => {
          console.log(`[${Strings.ProjectName}] Installation complete. Rendering web part.`);
          this._renderTracker();
        });
      } catch (modalErr) {
        console.error(`[${Strings.ProjectName}] Installation error:`, formatError(modalErr));
        this._renderError(formatError(modalErr));
      }

    }

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

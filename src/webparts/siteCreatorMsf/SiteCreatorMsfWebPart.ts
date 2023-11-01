import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SiteCreatorMsfWebPartStrings';
import SiteCreatorMsf from './components/SiteCreatorMsf';
import { ISiteCreatorMsfProps } from './components/ISiteCreatorMsfProps';

//Graph toolkit
import { Providers } from '@microsoft/mgt-element/dist/es6/providers/Providers';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider/dist/es6/SharePointProvider';

export interface ISiteCreatorMsfWebPartProps {
  expanded: boolean
}

export default class SiteCreatorMsfWebPart extends BaseClientSideWebPart<ISiteCreatorMsfWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISiteCreatorMsfProps> = React.createElement(
      SiteCreatorMsf,
      {
        context: this.context,
        expanded: this.properties.expanded
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected async onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    
    //return super.onInit();
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "General",
              groupFields: [
                PropertyPaneToggle('expanded',{
                  label: "Collapsed/Expanded display",
                  offText: "Collapsed",
                  onText: "Expanded"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

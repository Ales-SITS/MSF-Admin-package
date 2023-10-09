import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SiteCreatorMsfWebPartStrings';
import SiteCreatorMsf from './components/SiteCreatorMsf';
import { ISiteCreatorMsfProps } from './components/ISiteCreatorMsfProps';

//Graph toolkit
import { Providers } from '@microsoft/mgt-element/dist/es6/providers/Providers';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider/dist/es6/SharePointProvider';
//import { lazyLoadComponent } from '@microsoft/mgt-spfx-utils';

export interface ISiteCreatorMsfWebPartProps {
  description: string;
}

export default class SiteCreatorMsfWebPart extends BaseClientSideWebPart<ISiteCreatorMsfWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISiteCreatorMsfProps> = React.createElement(
      SiteCreatorMsf,
      {
        context: this.context,
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

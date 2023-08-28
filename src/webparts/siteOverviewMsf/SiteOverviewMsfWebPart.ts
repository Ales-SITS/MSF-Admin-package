import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SiteOverviewMsfWebPartStrings';
import ClassWrapper from './components/ClassWrapper';
import { ISiteOverviewMsfProps } from './components/ISiteOverviewMsfProps';

//GRAPH
import { SPFx, graphfi } from "@pnp/graph";

export interface ISiteOverviewMsfWebPartProps {
  header: string;
  site_id: string;
  site_url: string;
  expanded: boolean
}

export default class SiteOverviewMsfWebPart extends BaseClientSideWebPart<ISiteOverviewMsfWebPartProps> {

  private siteID:string = ""

  public async onInit(): Promise<void> {
    try {  
      this.siteID = await this.getID()     
      return super.onInit();
    } catch (error) {
      console.error('Error in onInit:', error);
      return super.onInit();
    }
  }
  

  public render(): void {
    const element: React.ReactElement<ISiteOverviewMsfProps> = React.createElement(
      ClassWrapper,
      {
        header: this.properties.header,
        site_id: this.siteID,
        site_url: this.properties.site_url,
        expanded: this.properties.expanded,
        context: this.context
        }
    );

    ReactDom.render(element, this.domElement);
  }

  
  private async getID() {
    const urlObject = new URL(this.properties.site_url);
    const host = urlObject.hostname
    const path = urlObject.pathname;
    const graph = graphfi().using(SPFx(this.context))
    const idstring = await graph.sites.getByUrl(host,path)()
    const id = idstring.id.split(",")[1]
  
    this.siteID = id

    return id
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath === 'site_url') {
      try {  
        this.siteID = await this.getID()           
      } catch (error) {
        console.error('Change:', error);   
      }

      this.render();
    }
  
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
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
              groupName: 'General',
              groupFields: [

                PropertyPaneTextField('site_url', {
                  label: "Site url"
                }),
                PropertyPaneLabel('site_url', {
                  text: `ID: ${this.siteID}`
                })
              ]
            },
            {
              groupName: 'Visual',
              groupFields: [
                PropertyPaneTextField('header', {
                  label: 'Header'
                }),             
                PropertyPaneToggle('expanded', {
                  label: "Fields Expanded/Collapsed?",
                  offText: "Expanded",
                  onText: "Collapsed"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

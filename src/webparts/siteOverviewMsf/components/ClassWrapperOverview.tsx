import * as React from 'react';
import { ISiteOverviewMsfProps } from './ISiteOverviewMsfProps';
import SiteOverviewMsf from './SiteOverviewMsf'
import { MSGraphClientV3 } from '@microsoft/sp-http';
//import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { spfi, SPFx as SPFxsp} from "@pnp/sp";


export default class ClassWrapper extends React.Component<ISiteOverviewMsfProps, {}> {
  public render(): React.ReactElement<ISiteOverviewMsfProps> {
    /*
    this.props.context.msGraphClientFactory
    .getClient('3')
    .then((client: MSGraphClientV3): void => {
    
      client
        .api(`/reports/getSharePointSiteUsageDetail(period='D7')`)
        //.api(`/me`)
        .get((error: any, response: any) => {
          console.log(error)
      });
    }
    )
*/
    const sp = spfi().using(SPFxsp(this.props.context))

    return (
      <>
        <SiteOverviewMsf details={this.props} sp={sp}/>
      </>
    );
  }
}

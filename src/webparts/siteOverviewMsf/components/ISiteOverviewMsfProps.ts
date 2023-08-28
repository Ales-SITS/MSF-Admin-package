import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISiteOverviewMsfProps {
  header: string;
  site_id: string;
  site_url: string;
  expanded: boolean;
  context: WebPartContext
}

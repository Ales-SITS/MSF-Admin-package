import * as React from 'react';
import styles from './SiteOverviewMsf.module.scss';
import { ISiteOverviewMsfProps } from './ISiteOverviewMsfProps';
import SiteOverviewMsf from './SiteOverviewMsf'


export default class ClassWrapper extends React.Component<ISiteOverviewMsfProps, {}> {
  public render(): React.ReactElement<ISiteOverviewMsfProps> {

    return (
      <>
        <SiteOverviewMsf details={this.props}/>
      </>
    );
  }
}

import * as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';
import { ISiteCreatorMsfProps } from './ISiteCreatorMsfProps';
import {  PeoplePicker } from '@microsoft/mgt-react';

//API
import { MSGraphClientV3  } from "@microsoft/sp-http";

//PNP SP
import { spfi, SPFx as SPFxsp} from "@pnp/sp";
import "@pnp/sp/site-designs"
import "@pnp/sp/sites";

import { IHubSiteInfo } from  "@pnp/sp/hubsites";
import "@pnp/sp/hubsites";

//PNP GRAPH
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";


export default function Declined (props) { 

  const user = props.context.pageContext.user.email.toLowerCase()

  return (
    <div className={styles.decline_wrapper}>
      {props.loader ? 
      <span className={styles.decline_loading}>Checking your permissions ...</span> :
      <span className={styles.decline_msg}>Your account {user} doesn't have permissions to create a site.</span>}
    </div>
  )
}

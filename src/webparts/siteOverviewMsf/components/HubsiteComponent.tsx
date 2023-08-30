import * as React from 'react';
import {useState, useEffect, useRef} from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPFx, graphfi } from "@pnp/graph";

import styles from './SiteOverviewMsf.module.scss';
import substyles from './SubsiteComponent.module.scss';

import { spfi, SPFx as SPFxsp} from "@pnp/sp";

import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";

import "@pnp/graph/users";
import "@pnp/graph/sites";



import { Icon } from '@fluentui/react/lib/Icon';
import { Site } from "@pnp/graph/sites";

export default function SubsiteComponent (props) {

    const hubsite = props.hubsite

    return (
        <div>
              <div className={styles.siteBoxTop}>
                    <div><a href={hubsite.SPWebUrl} title={hubsite.SPWebUrl}>{hubsite.Title}</a></div>
                    <div className={substyles.subSiteBoxBottom}>
                      <a href={`${hubsite.SPWebUrl}/_layouts/15/settings.aspx`} title="Site Settings"><Icon iconName="Settings"/></a>
                      <a href={`${hubsite.SPWebUrl}/_layouts/15/user.aspx`} title="Site Permissions"><Icon iconName="SecurityGroup"/></a>
                      <a href={`${hubsite.SPWebUrl}/_layouts/15/viewlsts.aspx?view=14`} title="Site Content"><Icon iconName="AllApps"/></a> 
                      <a href={`${hubsite.SPWebUrl}/_layouts/15/siteanalytics.aspx?view=19`} title="Site Usage"><Icon iconName="LineChart"/></a> 
                      <a href={`${hubsite.SPWebUrl}/_layouts/15/storman.aspx`} title="Site Storage"><Icon iconName="OfflineStorage"/></a>  
                    </div>
                    <div><span className={styles.idBox}>{hubsite.IdentitySiteCollectionId}</span></div>    
              </div>   
        </div>

    );
}

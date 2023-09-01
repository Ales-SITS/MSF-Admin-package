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


export default function LibraryComponent(props) {

    const list = props.list
    const siteurl = props.siteurl
    const context = props.context

     return (

        <div>
            <div className={styles.itemBoxTop}>
                <div><a href={list.url} title={list.url}>{list.name}</a></div>
                <div className={substyles.subsiteBoxBottom}>
                    <a className={`${styles.buttonMedium} ${styles.buttonMediumLibrary}`}  href={`${siteurl}/_layouts/15/listedit.aspx?List={${list.id}}`} title="Library Settings"><Icon iconName="Settings"/></a>
                    <a className={`${styles.buttonMedium} ${styles.buttonMediumLibrary}`}  href={`${siteurl}/_layouts/15/user.aspx?obj={${list.id}},doclib&List={${list.id}}`} title="Library Permissions"><Icon iconName="SecurityGroup"/></a>
                    <a className={`${styles.buttonMedium} ${styles.buttonMediumLibrary}`} href={`${siteurl}/_layouts/15/storman.aspx?root=${list.url.split("/")[3]}`} title="List Storage"><Icon iconName="OfflineStorage"/></a>  
                </div>
                <div><span className={styles.idBox}>{list.id}</span></div>    
            </div>   
        </div> 
    );
  }

  
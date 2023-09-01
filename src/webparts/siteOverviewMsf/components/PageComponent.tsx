import * as React from 'react';

import styles from './SiteOverviewMsf.module.scss';
import substyles from './SubsiteComponent.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';


export default function ListComponent (props) {

    const page = props.page
    const sitePages = props.sitePages
    const siteurl = props.siteurl
    const context = props.context

     return (

        <div>
            <div className={styles.itemBoxTop}>
                <div><a href={`${siteurl}/SitePages/${page.FileLeafRef}`} title={page.Title}>{page.Title}</a></div>
                    <div className={substyles.subsiteBoxBottom}>
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumPage}`} href={`${siteurl}/_layouts/15/user.aspx?obj={${sitePages.id}},doclib&List={${sitePages.id}}`} title="Page Permissions"><Icon iconName="SecurityGroup"/></a>
                    </div>
                <div><span className={styles.idBox}>{page.Guid}</span></div>    
            </div>   
        </div> 
    );
  }

import * as React from 'react';
import {useState, useEffect, useRef} from 'react';

import styles from './SiteOverviewMsf.module.scss';
import substyles from './SubsiteComponent.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';


export default function ListComponent (props) {

    const page = props.page
    const sitePages = props.sitePages
    const siteurl = props.siteurl
    const context = props.context


    const [idhidden,setIdhidden] = useState(true)
    const idHandler = (e) => {
      e.stopPropagation()
      setIdhidden(!idhidden)
    }  

  //OTHER
    const copyOnClick = (e) => {
        e.stopPropagation()
        navigator.clipboard.writeText(e.target.innerText)
    }

    console.log(page)

     return (
        <div className={styles.itemBoxWrapper} >
            <div className={styles.itemBox} >
            <div className={styles.itemBoxTop}>
                 <div className={styles.itemBoxTopLeft}>
                      <div className={styles.idBoxWrapper}>
                          <button className={styles.idBoxLabel} onClick={(e)=>idHandler(e)}>id</button>
                          {!idhidden && <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{page.GUID}</span>}   
                      </div>
                      <a className={styles.itemBoxSiteLink} href={`${siteurl}/SitePages/${page.FileLeafRef}`} title={`Go to ${siteurl}/SitePages/${page.FileLeafRef}`}>{page.Title}</a>
                  </div>

                  <div className={styles.itemBoxTopRight}>
                    <div className={`${styles.buttonBox} ${styles.buttonBoxPage}`}>
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumPage}`} href={`${siteurl}/_layouts/15/user.aspx?obj={${sitePages.id}},doclib&List={${sitePages.id}}`} title="Page Permissions"><Icon iconName="SecurityGroup"/></a>
                    </div>
                  </div>   
            </div>  
            </div> 
        </div> 
    );
  }

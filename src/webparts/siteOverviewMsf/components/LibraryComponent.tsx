import * as React from 'react';
import {useState} from 'react';

import styles from './SiteOverviewMsf.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';

export default function LibraryComponent(props) {

    const list = props.list
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

//URL
    const urlObject = new URL(siteurl);
    const host = urlObject.hostname

     return (
        <div className={styles.itemBoxWrapper} >
        <div className={styles.itemBox} >
        <div className={styles.itemBoxTop}>
             <div className={styles.itemBoxTopLeft}>
                  <div className={styles.idBoxWrapper}>
                      <button className={styles.idBoxLabel} onClick={(e)=>idHandler(e)}>id</button>
                      {!idhidden && <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{list.id}</span>}   
                  </div>
                  <a className={styles.itemBoxSiteLink} href={`https://${host}/${list.url}`} title={`Go to https://${host}/${list.url}`}>{list.name}</a>
              </div>
              <div className={styles.itemBoxTopRight}>
                <div className={`${styles.buttonBox} ${styles.buttonBoxLibrary}`}>
                    <a className={`${styles.buttonMedium} ${styles.buttonLibrary}`}  href={`${siteurl}/_layouts/15/listedit.aspx?List={${list.id}}`} title="Library Settings"><Icon iconName="Settings"/></a>
                    <a className={`${styles.buttonMedium} ${styles.buttonLibrary}`}  href={`${siteurl}/_layouts/15/user.aspx?obj={${list.id}},doclib&List={${list.id}}`} title="Library Permissions"><Icon iconName="SecurityGroup"/></a>
                    <a className={`${styles.buttonMedium} ${styles.buttonLibrary}`} href={`${siteurl}/_layouts/15/storman.aspx?root=${list.url.split("/")[3]}`} title="Library Storage"><Icon iconName="OfflineStorage"/></a>
                    
                   </div>
              </div>   
        </div>  
        </div> 
      
    </div> 
    );
  }

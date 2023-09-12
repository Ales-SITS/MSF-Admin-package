import * as React from 'react';
import {useState} from 'react';

//STYLING
import styles from './SiteOverviewMsf.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';

export default function SubsiteComponentDetails (props) {

    const site = props.site
    const libraries = props.libraries
    const lists = props.lists

//HIDERS
    const [idhidden,setIdhidden] = useState(true)
    const idHandler = (e):void => {
        e.stopPropagation()
        setIdhidden(!idhidden)
    }  

//OTHER
    const copyOnClick = (e):void => {
        e.stopPropagation()
        navigator.clipboard.writeText(e.target.innerText)
    }

     return (  
                  <div>
                      <span>Libraries</span>
                      <ul>
                        {
                          libraries.map((lib,idx)=>
                            <li key={idx} className={`${styles.itemBoxTop} ${styles.itemBoxTopSmall}`}>
                                    <div className={styles.itemBoxTopLeft}>
                                        <div className={styles.idBoxWrapper}>
                                            <button className={styles.idBoxLabel} onClick={(e)=>idHandler(e)}>id</button>
                                            {!idhidden && <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{lib.id}</span>}   
                                        </div>
                                        <a className={`${styles.itemBoxSiteLink} ${styles.itemBoxSiteLinkSmall}`} href={`${site}${lib.url}`} title={`Go to ${site}${lib.url}`}>{lib.name}</a>
                                    </div>
                                    <div className={styles.itemBoxTopRight}>
                                        <div className={`${styles.buttonBox} ${styles.buttonBoxLibrary}`}>
                                            <a className={`${styles.buttonSmall} ${styles.buttonLibrary}`} href={`${site}/_layouts/15/listedit.aspx?List={${lib.id}}`} title="Library Settings"><Icon iconName="Settings"/></a>
                                            <a className={`${styles.buttonSmall} ${styles.buttonLibrary}`} href={`${site}/_layouts/15/user.aspx?obj={${lib.id}},doclib&List={${lib.id}}`} title="Library Permissions"><Icon iconName="SecurityGroup"/></a>
                                            <a className={`${styles.buttonSmall} ${styles.buttonLibrary}`} href={`${site}/_layouts/15/storman.aspx?root=${lib.url.split("/")[3]}`} title="Library Storage"><Icon iconName="OfflineStorage"/></a>
                                        </div>
                                    </div>        
                                </li>                    
                        )}                  
                      </ul>
                      <span>Lists</span>
                      <ul>
                        {
                          lists.map((list,idx)=>
                          <li key={idx} className={`${styles.itemBoxTop} ${styles.itemBoxTopSmall}`}>
                          <div className={styles.itemBoxTopLeft}>
                              <div className={styles.idBoxWrapper}>
                                  <button className={styles.idBoxLabel} onClick={(e)=>idHandler(e)}>id</button>
                                  {!idhidden && <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{list.id}</span>}   
                              </div>
                              <a className={`${styles.itemBoxSiteLink} ${styles.itemBoxSiteLinkSmall}`} href={`${site}${list.url}`} title={`Go to ${site}${list.url}`}>{list.name}</a>
                          </div>
                          <div className={styles.itemBoxTopRight}>
                              <div className={`${styles.buttonBox} ${styles.buttonBoxList}`}>
                                  <a className={`${styles.buttonSmall} ${styles.buttonList}`} href={`${site}/_layouts/15/listedit.aspx?List={${list.id}}`} title="Library Settings"><Icon iconName="Settings"/></a>
                                  <a className={`${styles.buttonSmall} ${styles.buttonList}`} href={`${site}/_layouts/15/user.aspx?obj={${list.id}},doclib&List={${list.id}}`} title="Library Permissions"><Icon iconName="SecurityGroup"/></a>
                                  <a className={`${styles.buttonSmall} ${styles.buttonList}`} href={`${site}/_layouts/15/storman.aspx?root=${list.url.split("/")[3]}`} title="Library Storage"><Icon iconName="OfflineStorage"/></a>
                              </div>
                          </div>        
                      </li>                         
                        )}                  
                      </ul>   
                    </div>
    );
  }
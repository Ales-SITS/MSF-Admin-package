import * as React from 'react';
import {useState} from 'react';

import styles from './SiteOverviewMsf.module.scss';
import substyles from './SubsiteComponent.module.scss';

import { spfi, SPFx as SPFxsp} from "@pnp/sp";

import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";

import "@pnp/graph/users";
import "@pnp/graph/sites";

import HubsiteComponentDetails from "./HubsiteComponentDetails"
import { Web } from "@pnp/sp/webs";  

import { Icon } from '@fluentui/react/lib/Icon';

export default function SubsiteComponent (props) {

    const hubsite = props.hubsite

    const site_id = hubsite.SiteId
    const context = props.context

    const [listsLoading,setListsLoading] = useState(false)

    const [lists,setLists] = useState([])

    const [listsHidden, setListsHidden] = useState(true)

//DATA
async function getSubsiteLists():Promise<any> {
  setListsLoading(true);
  const sp = spfi().using(SPFxsp(context));
  const subhubsite = Web([sp.web, `${hubsite.Path}`]);
  const lists = await subhubsite.lists();
  setListsLoading(false);

  const listArr = [];

  const listPromises = lists.map((list) =>
    subhubsite.lists
      .getById(list.Id)
      .select("Title", "RootFolder/ServerRelativeUrl")
      .expand("RootFolder")()
      .then((response) => {
        listArr.push({
          name: list.Title,
          id: list.Id,
          url: response.RootFolder.ServerRelativeUrl,
          template: list.BaseTemplate,
        });
      })
  );

  return Promise.all(listPromises).then(() => {
    return listArr;
  });
}

//HIDDERS
const listHandler = ():void => {
    setListsHidden(!listsHidden)
    getSubsiteLists().then(result => {

      setLists([]);
      const arr:any = result

      setLists(arr);
    })   
  }

const [idhidden,setIdhidden] = useState(true)
const idHandler = (e) => {
      e.stopPropagation()
      setIdhidden(!idhidden)
  }  

//URL
    const urlObject = new URL(hubsite.Path);
    const host = urlObject.hostname

//FILTERED VALUES
      const filteredLists = lists.filter( list => list.template === 100 && list.name!=="DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES" && !list.name.includes("@thread.skype_wiki"))
      const filteredLibs = lists.filter( lib => lib.template === 101)

//OTHER
      const copyOnClick = (e):void => {
          e.stopPropagation()
          navigator.clipboard.writeText(e.target.innerText)
          }
      //console.log(hubsite)
    return (
        <div className={styles.itemBoxWrapper} onClick={listHandler}>
            <button className={listsHidden ? styles.arrow : `${styles.arrow} ${styles.arrowOpened}`}>▶</button>
            <div className={styles.itemBox} >
              <div className={styles.itemBoxTop} >
                  <div className={styles.itemBoxTopLeft}>
                      <div className={styles.idBoxWrapper}>
                          <button className={styles.idBoxLabel} onClick={(e)=>idHandler(e)}>id</button>
                          {!idhidden && <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{site_id}</span>}   
                      </div>
                      <a className={styles.itemBoxSiteLink} href={hubsite.SPWebUrl} title={`Go to ${hubsite.SPWebUrl}`}>{hubsite.Title}</a>
                  </div>
                  <div className={styles.itemBoxTopRight}>
                    <div className={`${styles.buttonBox} ${styles.buttonBoxHubsite}`}>
                      <a className={`${styles.buttonMedium} ${styles.buttonMediumSite}`} href={`${hubsite.SPWebUrl}/_layouts/15/viewlsts.aspx?view=14`} title="Site Content"><Icon iconName="AllApps"/></a> 
                      <a className={`${styles.buttonMedium} ${styles.buttonMediumSite}`} href={`${hubsite.SPWebUrl}/_layouts/15/settings.aspx`} title="Site Settings"><Icon iconName="Settings"/></a>
                      <a className={`${styles.buttonMedium} ${styles.buttonMediumSite}`} href={`${hubsite.SPWebUrl}/_layouts/15/user.aspx`} title="Site Permissions"><Icon iconName="SecurityGroup"/></a>
                      <a className={`${styles.buttonMedium} ${styles.buttonMediumSite}`} href={`${hubsite.SPWebUrl}/_layouts/15/siteanalytics.aspx?view=19`} title="Site Usage"><Icon iconName="LineChart"/></a> 
                      <a className={`${styles.buttonMedium} ${styles.buttonMediumSite}`} href={`${hubsite.SPWebUrl}/_layouts/15/storman.aspx`} title="Site Storage"><Icon iconName="OfflineStorage"/></a>  
                    </div>
                  </div>
              </div>   
              {!listsHidden &&
                <div className={styles.itemBoxBottom}>
                {listsLoading? 
                <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div>:     
                <HubsiteComponentDetails libraries={filteredLibs} lists={filteredLists} site={hubsite.Path}/>
                }   
            </div>
              }
              </div>            
        </div>

    );
}



/*
         
              {!listsHidden &&
                <div className={styles.itemBoxBottom}>
                  {listsLoading? 
                  <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div>:     
                  <SubsiteComponentDetails libraries={filteredLibs} lists={filteredLists} site={site.Url}/>
                  }   
              </div>
              }
*/


/*
<div className={styles.itemBoxBottom}>
                  {listsLoading? 
                  <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div>:     
                    <div>
                      <span>Libraries</span>
                      <ul>
                        {
                          filteredLibs.map((lib,idx)=>
                            <li key={idx} className={substyles.subsite}>
                              <div><a href={`https://${host}${lib.url}`}>{lib.name}</a></div>
                              <div className={substyles.subsiteBottom}>
                                  <a className={styles.buttonMedium} href={`${hubsite.Path}/_layouts/15/listedit.aspx?List={${lib.id}}`} title="Library Settings"><Icon iconName="Settings"/></a>
                                  <a className={styles.buttonMedium} href={`${hubsite.Path}/_layouts/15/user.aspx?obj={${lib.id}},doclib&List={${lib.id}}`} title="Library Permissions"><Icon iconName="SecurityGroup"/></a>
                                  <a className={styles.buttonMedium} href={`${hubsite.Path}/_layouts/15/storman.aspx?root=${lib.url.split("/")[3]}`} title="List Storage"><Icon iconName="OfflineStorage"/></a>  
                              </div>
                              <div onClick={(e)=>copyOnClick(e)}><span className={styles.idBox}>{lib.id}</span></div>    
                            </li>                    
                        )}                  
                      </ul>
                      <span>Lists</span>
                      <ul>
                        {
                          filteredLists.map((list,idx)=>
                            <li key={idx} className={substyles.subsite}>
                              <div><a href={`https://${host}${list.url}`}>{list.name}</a></div>
                              <div className={substyles.subsiteBottom}>
                                <a className={styles.buttonMedium} href={`${hubsite.Path}/_layouts/15/listedit.aspx?List={${list.id}}`} title="Subsite Settings"><Icon iconName="Settings"/></a>
                                <a className={styles.buttonMedium} href={`${hubsite.Path}/_layouts/15/user.aspx?obj={${list.id}},doclib&List={${list.id}}`} title="List Permissions"><Icon iconName="SecurityGroup"/></a>
                              </div>
                              <div onClick={(e)=>copyOnClick(e)}><span className={styles.idBox}>{list.id}</span></div>    
                            </li>                    
                        )}                  
                      </ul>
                    </div>
                  }   
              </div>


*/
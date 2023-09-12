import * as React from 'react';
import {useState} from 'react';

//STYLING
import styles from './SiteOverviewMsf.module.scss';
import substyles from './SubsiteComponent.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';

//API
//import { SPFx, graphfi } from "@pnp/graph";
import { spfi, SPFx as SPFxsp} from "@pnp/sp";


import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

import "@pnp/sp/site-groups";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/search";

import "@pnp/graph/users";
import "@pnp/graph/sites";
import "@pnp/sp/hubsites";
import "@pnp/graph/search";

import { Web } from "@pnp/sp/webs";  

//Components

import SubsiteComponentDetails from "./SubsiteComponentDetails"


export default function SubsiteComponent (props) {

    const site = props.site
    const site_id = site.Id
    const context = props.context

    const [listsLoading,setListsLoading] = useState(false)

    const [lists,setLists] = useState([])

    const [listsHidden, setListsHidden] = useState(true)

//DATA
async function getSubsiteLists(id):Promise<any> {
  setListsLoading(true);
  const sp = spfi().using(SPFxsp(context));
  const subsite = Web([sp.web, `${site.Url}`]);
  const lists = await subsite.lists();
  setListsLoading(false);

  const listArr = [];

  const listPromises = lists.map((list) =>
    subsite.lists
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

//HIDERS
  const listHandler = ():void => {
    setListsHidden(!listsHidden)
    getSubsiteLists(site_id).then(result => {

      setLists([]);
      const arr:any = result

      setLists(arr);
    })   
  }
  const [idhidden,setIdhidden] = useState(true)
  const idHandler = (e):void => {
      e.stopPropagation()
      setIdhidden(!idhidden)
  }  

//FILTERED VALUES
      const filteredLists = lists.filter( list => list.template === 100 && list.name!=="DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES")
      const filteredLibs = lists.filter( lib => lib.template === 101)

//OTHER
      const copyOnClick = (e):void => {
        e.stopPropagation()
          navigator.clipboard.writeText(e.target.innerText)
        }

     return (
        <div className={styles.itemBoxWrapper} onClick={listHandler}>
            <button className={listsHidden ? styles.arrow : `${styles.arrow} ${styles.arrowOpened}`}>▶</button>
            <div className={styles.itemBox}>
              <div className={listsHidden ? styles.itemBoxTop : `${styles.itemBoxTop} ${styles.itemBoxTopSelected}`  }>
                    <div className={styles.itemBoxTopLeft}>
                      <div className={styles.idBoxWrapper}>
                          <button className={styles.idBoxLabel} onClick={(e)=>idHandler(e)}>id</button>
                          {!idhidden && <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{site_id}</span>}   
                      </div>
                      <a className={styles.itemBoxSiteLink} href={site.Url} title={`Go to ${site.Url}`}>{site.Title}</a>
                    </div>
                    <div className={styles.itemBoxTopRight}>
                      <div className={`${styles.buttonBox} ${styles.buttonBoxSubsite}`}>
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumSubsite}`} href={`${site.Url}/_layouts/15/viewlsts.aspx?view=14`} title="Subsite Content"><Icon iconName="AllApps"/></a> 
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumSubsite}`} href={`${site.Url}/_layouts/15/settings.aspx`} title="Subsite Settings"><Icon iconName="Settings"/></a>
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumSubsite}`} href={`${site.Url}/_layouts/15/user.aspx`} title="Subsite Permissions"><Icon iconName="SecurityGroup"/></a>
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumSubsite}`} href={`${site.Url}/_layouts/15/siteanalytics.aspx?view=19`} title="Subsite Usage"><Icon iconName="LineChart"/></a> 
                        <a className={`${styles.buttonMedium} ${styles.buttonMediumSubsite}`} href={`${site.Url}/_layouts/15/storman.aspx`} title="Site Storage"><Icon iconName="OfflineStorage"/></a>  
                        <div className={`${styles.buttonModernToClassic}`}>
                            <a className={`${styles.buttonMedium} ${styles.buttonMediumSubsite} ${styles.buttonModern}`} href={`${site.Url}/_layouts/15/AdminRecycleBin.aspx`} title="Site Recycle Bin"><Icon iconName="RecycleBin"/></a>
                            <a className={`${styles.buttonMedium} ${styles.buttonClassic}`} href={`${site.Url}/_layouts/15/AdminRecycleBin.aspx?View=2`} title="2nd stage Site Recycle Bin"><Icon iconName="EmptyRecycleBin"/></a> 
                        </div>
                      
                      </div>
                    </div>
              </div>           
              {!listsHidden &&
                <div className={styles.itemBoxBottom}>
                  {listsLoading? 
                  <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div>:     
                  <SubsiteComponentDetails libraries={filteredLibs} lists={filteredLists} site={site.Url}/>
                  }   
              </div>
              }
            </div>
        </div>
    );
  }



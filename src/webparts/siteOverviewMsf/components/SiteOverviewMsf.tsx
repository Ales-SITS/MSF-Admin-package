import * as React from 'react';
import {useState, useEffect, useRef} from 'react';

//STYLING
import styles from './SiteOverviewMsf.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

//API
import { SPFx, graphfi } from "@pnp/graph";
import { spfi, SPFx as SPFxsp} from "@pnp/sp";

import "@pnp/sp/sites";
import "@pnp/sp/webs";

import "@pnp/sp/site-groups";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/search";

import "@pnp/graph/users";
import "@pnp/graph/sites";
import "@pnp/sp/hubsites";
import "@pnp/graph/search";

import { Web } from "@pnp/sp/webs";   

import { SearchResults } from "@pnp/sp/search";
import { Site } from "@pnp/graph/sites";

//COMPONENTS
import HubsiteComponent from "./HubsiteComponent"
import SubsiteComponent from "./SubsiteComponent"
import ListComponent from "./ListComponent"
import PermissionsComponent from "./PermissionsComponent"
import PnP_Generator from "../PnPScripts/PnP_Generator"

export default function SiteOverviewMsf (props) {
    const {
        header,
        site_id,
        site_url,
        expanded,
        context
      } = props.details;

    const [hubLoading,setHubLoading] = useState(true)
    const [subLoading,setSubLoading] = useState(true)
    const [listsLoading,setListsLoading] = useState(true)

    const [siteTitle,setSiteTitle] = useState(context.pageContext.web.title)

    const [siteURL,setSiteURL] = useState(
        site_id === undefined || site_id === null || site_id === "" ? context.pageContext.site.absoluteUrl : site_url)
    const [siteID,setSiteID] = useState(
        site_id === undefined || site_id === null || site_id === "" ? context.pageContext.site.id._guid : site_id)

    const [hubsites, setHubsites] = useState([])
    const [hubhide, setHubhide] = useState(expanded)
    const hubhideHandler = () => {
        setHubhide(!hubhide)
    }

    const [subsites, setSubsites] = useState([]);
    const [subhide, setSubhide] = useState(expanded)
    const subhideHandler = () => {
        setSubhide(!subhide)
    }

    const [lists,setLists] = useState([])
    const [libhide,setLibhide] = useState(expanded)
    const libhideHandler = () => {
        setLibhide(!libhide)
    }

    const [lishide,setLishide] = useState(expanded)
    const lishideHandler = () => {
        setLishide(!lishide)
    }

    async function getHub(id) {
        setHubLoading(true)
        const sp = spfi().using(SPFxsp(context));
        const searchResults: SearchResults = await sp.search(
          `DepartmentId=${id} contentclass:sts_site -SiteId:${id}`
        );     
        const result = searchResults.PrimarySearchResults
        setHubLoading(false)
        return result
    }  

    async function getSubsites(id) {   
        setSubLoading(true)
        const sp = spfi().using(SPFxsp(context));     
        const site = Web([sp.web, `${siteURL}`])      
        const sites = await site.webs()    
        setSubLoading(false)

        return sites
    }   

    async function getSiteCollectionLists(id) {
        setListsLoading(true)
        const graph = graphfi().using(SPFx(context))
        const siteData = await graph.sites.getById(id)
        const lists = await Site(siteData, "lists")();
        setListsLoading(false)
        return lists
    }   
  

    //TESTING
    async function getSubsiteLists(id) {
        setListsLoading(true);
        const sp = spfi().using(SPFxsp(context));
        const subsite = Web([sp.web, `${site_url}`]);
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
     //TESTING


    useEffect(() => {
        getSubsites(siteID).then(result => {
            setSubsites([]);
            const arr:any = result
            setSubsites(arr);
        });

        getHub(siteID).then(result => {
            setHubsites([]);
            const arr:any = result
            setHubsites(arr);
        });

        getSiteCollectionLists(siteID).then(result => {
            setLists([]);
            const arr:any = result
            setLists(arr);
        });

        getSubsiteLists(site_id).then(result => {     
            setLists([]);
            const arr:any = result
            setLists(arr);
          })   

       }, [props.details, siteID]);

       const copyOnClick = (e) => {
        navigator.clipboard.writeText(e.target.innerText)
       }

    useEffect(()=>{
        site_id === undefined || site_id === null || site_id === "" ? 
        setSiteID(context.pageContext.site.id._guid) :
        setSiteID(site_id)
    },[site_id])

    useEffect(()=>{
        header === undefined || header === null || header === "" ? 
        setSiteTitle(context.pageContext.web.title) :
        setSiteTitle(header)
    },[header])

       //const libraries = lists.filter( lib => lib.list.template === "documentLibrary")
       //const genlist = lists.filter( lib => lib.list.template === "genericList" && lib.displayName !== "DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES")


       const [subFilter,setSubFilter] = useState("")
       const searchFilter = (e) => {
           setSubFilter(e)
       }

       const subsitesfiltered = subsites.filter( sub => sub.Title.includes(subFilter))
    

       const genlist = lists.filter( list => list.template === 100 && list.name!=="DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES")
       const libraries = lists.filter( lib => lib.template === 101)

    const [permVis,setPermVis] = useState(false)
    const permVisHandler = () =>{
        setPnpVis(false)
        setPermVis(!permVis)
    }
 
    const [pnpVis,setPnpVis] = useState(false)
    const pnpVisHandler = () =>{
        setPermVis(false)
        setPnpVis(!pnpVis)
    }
 


     return (
        <div className={styles.overviewWrapper}>
            <div className={styles.mainSiteBox}>
                <div className={styles.mainSiteBoxTop}>
                    <h1><a href={siteURL} title={siteURL}>{siteTitle}</a></h1>
                    <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{siteID}</span>
                </div>
                <div className={styles.mainSiteBoxBottom}>
                    <a href={`${siteURL}/_layouts/15/viewlsts.aspx?view=14`} title="Site Content"><Icon iconName="AllApps"/></a> 
                    <a href={`${siteURL}/_layouts/15/settings.aspx`} title="Site Settings"><Icon iconName="Settings"/></a>
                    <a href={`${siteURL}/_layouts/15/user.aspx`} title="Site Permissions"><Icon iconName="SecurityGroup"/></a>
                    <a href={`${siteURL}/_layouts/15/siteanalytics.aspx?view=19`} title="Site Usage"><Icon iconName="LineChart"/></a>  
                    <a href={`${siteURL}/_layouts/15/storman.aspx`} title="Site Storage"><Icon iconName="OfflineStorage"/></a> 
                    <button onClick={pnpVisHandler} title="PnP Scripts" className="PnPScriptButton"><Icon iconName="PasteAsCode"/></button>
                    <button onClick={permVisHandler} title="Permissions"><Icon iconName="PeopleAlert"/></button> 
                </div>
            </div>
            {permVis && <PermissionsComponent onCloseHandler={permVisHandler} context={context} url={siteURL}/>}
            {pnpVis && <PnP_Generator onCloseHandler={pnpVisHandler} type={"top_site"} siteurl={site_url}/>}
            <div className={styles.detailsWrapper}>
                <button className={styles.detailsWrapperButton} onClick={hubhideHandler}>
                    <span>{hubhide ? "▲ " : "▼ "} Hub associated sites</span>
                    {
                    hubLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span>({hubsites.length})</span>
                     }
                </button>
                {!hubhide &&
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {hubsites.map((hubsite,idx)=>
                            <li key={idx}>
                                <HubsiteComponent hubsite={hubsite} context={context}/>
                            </li>
                            )}
                        </ul>
                    </div>
                }
            </div>
            <div className={styles.detailsWrapper}>
                <button className={styles.detailsWrapperButton} onClick={subhideHandler}>
                    <span>{subhide ? "▲ " : "▼ "} Subsites</span>
                    {
                    subLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span>({subsites.length})</span>
                    }
                    
                </button>
                {!subhide && 
                    <div className={styles.resultsWrapper}>
                        <input 
                        type="text" 
                        name="siteName" 
                        placeholder="Filter by site Title"
                        onChange={e => searchFilter(e.target.value)} 
                        /><span>({subsitesfiltered.length})</span>
                        <ul>
                            {subsitesfiltered.map((site,idx)=>
                            <li key={idx}>
                                <SubsiteComponent site={site} context={context}/>
                            </li>
                            )}
                        </ul>
                    </div>
                }
            </div>
            <div className={styles.detailsWrapper}>
                <button className={styles.detailsWrapperButton} onClick={libhideHandler}>
                    <span>{libhide ? "▲ " : "▼ "} Libraries</span>
                    {
                    listsLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span>({libraries.length})</span>
                    }
                </button>
                {!libhide && 
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {libraries.map((list,idx)=>
                                <li key={idx}>
                                    <ListComponent list={list} siteurl={siteURL}/>
                                </li>
                            )}
                        </ul>
                    </div>
                }
            </div>
            <div className={styles.detailsWrapper}>
                <button className={styles.detailsWrapperButton} onClick={lishideHandler}>
                    <span>{lishide ? "▲ " : "▼ "} Lists</span>
                    {
                    listsLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span>({genlist.length})</span>
                    }
                </button>
                {!lishide && 
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {genlist.map((list,idx)=>
                                <li key={idx}>
                                    <ListComponent list={list} siteurl={siteURL}/>
                                </li>
                            )}
                        </ul>
                    </div>
                }
            </div>
        </div>
    );
  }

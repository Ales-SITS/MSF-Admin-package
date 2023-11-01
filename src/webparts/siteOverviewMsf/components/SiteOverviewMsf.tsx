import * as React from 'react';
import {useState, useEffect, useRef} from 'react';

//STYLING
import styles from './SiteOverviewMsf.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

//API
//import { SPFx, graphfi} from "@pnp/graph";
//import { spfi, SPFx as SPFxsp} from "@pnp/sp";

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
import "@pnp/sp/features";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { IHubSiteInfo } from  "@pnp/sp/hubsites";
import "@pnp/sp-admin";

import { Web } from "@pnp/sp/webs";   

import { SPFx, graphfi } from "@pnp/graph";

import { SearchResults } from "@pnp/sp/search";
import { Site } from "@pnp/graph/sites";
import "@pnp/graph/groups";
import { GroupType } from '@pnp/graph/groups';
/*
import { MSGraphClient } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';  
*/

//COMPONENTS
import HubsiteComponent from "./HubsiteComponent"
import SubsiteComponent from "./SubsiteComponent"
import PageComponent from "./PageComponent"
import LibraryComponent from "./LibraryComponent"
import ListComponent from "./ListComponent"

import PermissionsComponent from "./PermissionsComponent"
import PnP_Generator from "../PnPScripts/PnP_Generator"


export default function SiteOverviewMsf (props) {
//PROPS
    const {
        header,
        site_id,
        site_url,
        expanded,
        dynamic_url,
        context
      } = props;

    const sp = spfi().using(SPFxsp(context))
    //const sp = props.sp//spfi().using(SPFxsp(context))

//LOADERS      
    const [hubLoading,setHubLoading] = useState(true)
    const [subLoading,setSubLoading] = useState(true)
    const [listsLoading,setListsLoading] = useState(true)
    const [pagesLoading,setPagesLoading] = useState(true)

 
//CONST
    const [siteTitle,setSiteTitle] = useState(context.pageContext.web.title)
      
    const siteURLfromProp = site_url === undefined || site_url === null || site_url === "" ? context.pageContext.site.absoluteUrl : site_url
    const [siteURL,setSiteUrl] = useState(siteURLfromProp)

    const [siteID,setSiteID] = useState(
        site_id === undefined || site_id === null || site_id === "" ? context.pageContext.site.id._guid : site_id)

    const [tabSelected,setTabSelected] = useState(1)
    const tabSelectedHandler = (tab) => {
        setTabSelected(tab)
    }
    
    const searchDynamicSite = (url):void => {
        url === "" ? setSiteUrl(siteURLfromProp) : setSiteUrl(url)
    }

//CONST & HIDERS 
    const [hubsites, setHubsites] = useState([])

    const [subsites, setSubsites] = useState([]);

    const [pages, setPages] = useState([])

    const [lists,setLists] = useState([])

    const [storage, setStorage] = useState(0)

    const [expand, setExpand] = useState(expanded)
    const expandHandler = () => {
        setExpand(!expand)
    }

    //GETTERS
    async function getSite():Promise<any> {  
        console.log(siteURL)
        console.log(siteURLfromProp)
        const site = Web([sp.web, `${siteURL}`])      
        const siteInfo = await site()           
        setSiteTitle(siteInfo.Title)

        try {
            const site_ID = await getSiteID()
            setSiteID(site_ID)
        } catch (error) {
            console.log(error)
        }

        try {
            const hub = await getHub(siteID)
            setHubsites(hub)
        } catch (error) {
            console.log(error)
        }

        try {
            const sub = await getSubsites()
            setSubsites(sub)
        } catch (error) {
            console.log(error)
        }

        try {      
        const pages = await getPages()
        setPages(pages)
        } catch (error) {
            console.log(error)
        }

        try {
            const lists = await getLists()
            setLists(lists)
        } catch (error) {
            console.log(error)
        }


        const storageReport = await getReport().then(result => {
            const parser = new DOMParser()
            const xml = parser.parseFromString(result,"application/xhtml+xml")
            const defaultNamespaceURI = 'http://schemas.microsoft.com/ado/2007/08/dataservices';
            const elementName = 'Storage';
            const storageElement = xml.getElementsByTagNameNS(defaultNamespaceURI, elementName)[0];         
            const storage = Number(storageElement.textContent) / (1024 * 1024 * 1024)

            return storage      
        });

        setStorage(storageReport)

    }

    async function getSiteID() {
        const urlObject = new URL(siteURL);
        const host = urlObject.hostname
        const path = urlObject.pathname;
        const graph = graphfi().using(SPFx(context))
        const idstring = await graph.sites.getByUrl(host,path)()
        const id = idstring.id.split(",")[1]   
        
        return id
    }

    async function getHub(id) {
        setHubLoading(true)
        const searchResults: SearchResults = await sp.search(
          `DepartmentId=${id} contentclass:sts_site -SiteId:${id}`
        );     
        const result = searchResults.PrimarySearchResults
        setHubLoading(false)
        return result
    }  

    async function getSubsites() {   
        setSubLoading(true)
        const site = Web([sp.web, `${siteURL}`])      
        const sites = await site.webs()    
        setSubLoading(false)

        return sites
    }   

    async function getAdvancedReport(id) {
        console.log("triggered")
        //const graph = graphfi().using(SPFx(context))
        //const siteData = await graph.sites.getById(id)
        //const report = await Site(siteData, "/reports/getSharePointSiteUsageStorage(period='D7')")();
        const site = Web([sp.web, `${siteURL}`])    
        //const features = await site.features()
        //const allProps = await sp.admin.siteProperties.select("*")();  
        const features = await site.features.select("DisplayName", "DefinitionId")();
        return features
    }  

    async function getReport():Promise<any> {
        const response = await fetch(`${siteURL}/_api/site/usage`);
        const rawData = await response.text()
        
        return rawData
    }   
  
    async function getPages():Promise<any> {   
        setPagesLoading(true)
        const site = Web([sp.web, `${siteURL}`])      
        const sites = await site.lists.getByTitle("Site Pages").items.select('FileLeafRef', 'Title', 'Id', 'GUID')()
        setPagesLoading(false)
        
        return sites
    }   

    async function getLists():Promise<any> {  
        setListsLoading(true);
        const subsite = Web([sp.web, `${siteURL}`]);
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

    useEffect(() => {  
        getSite()
       }, [siteID, siteURL]);

    useEffect(() => {
        setSiteUrl(siteURLfromProp)
    },[site_url])


    useEffect(()=>{
        header === undefined || header === null || header === "" ? 
        setSiteTitle(context.pageContext.web.title) :
        setSiteTitle(header)
    },[header])

//FILTERS    
    const genlist = lists.filter( list => list.template === 100 && list.name!=="DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES")
    const libraries = lists.filter( lib => lib.template === 101)
    const sitePages = lists.filter(lib => lib.template === 119 )   

    const [subFilter,setSubFilter] = useState("")
    const searchSubFilter = (e):void => {
           setSubFilter(e)
       }
    const subsitesfiltered = subFilter === "" ? subsites : subsites.filter( sub => sub.Title.toLowerCase().includes(subFilter.toLowerCase()))

    const [hubFilter,setHubFilter] = useState("")
    const searchHubFilter = (e):void => {
           setHubFilter(e)
       }    
    const hubsitesfiltered = hubFilter === "" ? hubsites : hubsites.filter( hub => hub.Title.toLowerCase().includes(hubFilter.toLowerCase()))

    
    const [pageFilter,setPageFilter] = useState("")
    const searchPageFilter = (e):void => {
        setPageFilter(e)
    }
    const pagesfiltered = pages.length === 0 ? [] : pageFilter === "" ? pages : pages.filter( page => page.Title.toLowerCase().includes(pageFilter.toLowerCase()))

    const [libFilter,setLibFilter] = useState("")
    const searchLibFilter = (e):void => {
        setLibFilter(e)
    }
    const libfiltered = libFilter === "" ? libraries : libraries.filter( lib => lib.name.includes(libFilter))

    const [listFilter,setListFilter] = useState("")
    const searchListFilter = (e):void => {
        setListFilter(e)
    }
    const listfiltered = listFilter === "" ? genlist : genlist.filter( list => list.name.includes(listFilter))

//MODALS   
    const [pnpVis,setPnpVis] = useState(false)
    const [permVis,setPermVis] = useState(false)
    const permVisHandler = ():void =>{
        setPnpVis(false)
        setPermVis(!permVis)
    }
 
    const pnpVisHandler = ():void =>{
        setPermVis(false)
        setPnpVis(!pnpVis)
    }

//UX
const copyOnClick = (e) => {
    navigator.clipboard.writeText(e.target.innerText)
   }

     return (
        <div className={styles.overviewWrapper}>
            <div className={styles.mainSiteBox}>
               {dynamic_url &&
                            <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="siteName" 
                            placeholder="Site URL"
                            onChange={e => searchDynamicSite(e.target.value)} 
                            />
                }
                <div className={styles.mainSiteBoxTop}>
                    <h1><a href={siteURL} title={siteURL}>{siteTitle}</a></h1>
                    <div>
                        <span className={styles.idBox} onClick={(e)=>copyOnClick(e)}>{siteID}</span>
                        <a href={siteURL} title={siteURL}>{siteURL}</a>
                        <span className={styles.storageWrapper}>storage used<span className={styles.storage}>{` ${storage.toFixed(3)} (GB)`}</span></span>
                    </div>
                </div>
                <div className={styles.mainSiteBoxBottom}>
                    <div className={styles.mainSiteBoxBottomLeft}>
                        <a className={styles.buttonModern} href={`${siteURL}/_layouts/15/viewlsts.aspx?view=14`} title="Site Content"><Icon iconName="AllApps"/></a>
                       
                        <a href={`${siteURL}/_layouts/15/settings.aspx`} title="Site Settings"><Icon iconName="Settings"/></a>
                        <a href={`${siteURL}/_layouts/15/user.aspx`} title="Site Permissions"><Icon iconName="SecurityGroup"/></a>
                        <div className={styles.mainSiteBoxBottomLeft}>
                            <a className={styles.buttonModern} href={`${siteURL}/_layouts/15/appStore.aspx`} title="App store"><Icon iconName="Puzzle"/></a>
                            <a className={styles.buttonClassic} href={`${siteURL}/AppCatalog/Forms/AllItems.aspx`} title="Site App cataloge"><Icon iconName="Puzzle"/></a>
                        </div>
                        <a href={`${siteURL}/_layouts/15/siteanalytics.aspx?view=19`} title="Site Usage"><Icon iconName="LineChart"/></a>  
                        <a href={`${siteURL}/_layouts/15/storman.aspx`} title="Site Storage"><Icon iconName="OfflineStorage"/></a> 
                        <div className={styles.mainSiteBoxBottomLeft}>
                            <a className={styles.buttonModern} href={`${siteURL}/_layouts/15/AdminRecycleBin.aspx?View=1`} title="Site Recycle Bin"><Icon iconName="RecycleBin"/></a>
                            <a className={styles.buttonClassic} href={`${siteURL}/_layouts/15/AdminRecycleBin.aspx?View=2`} title="2nd stage Site Recycle Bin"><Icon iconName="EmptyRecycleBin"/></a> 
                        </div>
                        <button onClick={pnpVisHandler} title="PnP Scripts App" className="PnPScriptButton"><Icon iconName="PasteAsCode"/></button>
                        <button onClick={permVisHandler} title="Permissions App"><Icon iconName="PeopleAlert"/></button>
                    </div>
                    <button onClick={expandHandler} title={`${expand ? "Click to collapse" : "Click to expand"}`} 
                                className={expand ? styles.mainSiteBoxBottomRight : `${styles.mainSiteBoxBottomRight} ${styles.mainSiteBoxBottomRightHidden} `}>
                                    ▲
                    </button>
                </div>
            </div>
            {permVis && <PermissionsComponent onCloseHandler={permVisHandler} context={context} url={siteURL} sp={sp}/>}
            {pnpVis && <PnP_Generator onCloseHandler={pnpVisHandler} type={"top_site"} siteurl={siteURL}/>}
            {expand && 
            <>
            <div className={styles.tabButtons}>
                <button onClick={()=>tabSelectedHandler(1)} className={tabSelected === 1 && styles.tabSelected}>
                    <span>Hub sites</span>
                    {hubLoading ? <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div> :
                    <span><span className={styles.displayedNum}>{hubsitesfiltered.length}/</span>{hubsites.length}</span>}
                </button>
                <button onClick={()=>tabSelectedHandler(2)} className={tabSelected === 2 && styles.tabSelected}>
                    <span>Subsites</span>
                    {subLoading ? <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div> :
                    <span><span className={styles.displayedNum}>{subsitesfiltered.length}/</span>{subsites.length}</span>}
                </button>
                <button onClick={()=>tabSelectedHandler(3)} className={tabSelected === 3 && styles.tabSelected}>
                    <span>Pages </span>
                    {pagesLoading ? <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div> :
                    <span><span className={styles.displayedNum}>{pagesfiltered.length}/</span>{pages.length}</span>}
                </button>
                <button onClick={()=>tabSelectedHandler(4)} className={tabSelected === 4 && styles.tabSelected}>
                    <span>Libraries</span> 
                    {listsLoading ? <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div> :
                    <span><span className={styles.displayedNum}>{libfiltered.length}/</span>{libraries.length}</span>}
                </button>
                <button onClick={()=>tabSelectedHandler(5)} className={tabSelected === 5 && styles.tabSelected}>
                    <span>Lists</span>
                    {listsLoading ? <div className={styles.loaderWrapper}><div className={styles.loader}><div></div><div></div><div></div><div></div></div></div> :
                    <span><span className={styles.displayedNum}>{listfiltered.length}/</span>{genlist.length}</span>}
                </button>
            </div>
            {tabSelected === 1 && 
            <div className={styles.detailsWrapper}>
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="siteName" 
                            placeholder="Filter by the site title"
                            onChange={e => searchHubFilter(e.target.value)} 
                            />
                    </div>
          
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {hubsitesfiltered.map((hubsite,idx)=>
                            <li key={idx}>
                                <HubsiteComponent hubsite={hubsite} context={context}/>
                            </li>
                            )}
                        </ul>
                    </div>    
            </div>
             }
            {tabSelected === 2 && 
            <div className={styles.detailsWrapper}>               
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="subsiteName" 
                            placeholder="Filter by the subsite title"
                            onChange={e => searchSubFilter(e.target.value)} 
                            />
                    </div>
                    <div className={styles.resultsWrapper}>

                        <ul>
                            {subsitesfiltered.map((site,idx)=>
                            <li key={idx}>
                                <SubsiteComponent site={site} context={context}/>
                            </li>
                            )}
                        </ul>
                    </div>
            </div>
            }
            {tabSelected === 3 && 
            <div className={styles.detailsWrapper}>              
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="pageName" 
                            placeholder="Filter by the page title"
                            onChange={e => searchPageFilter(e.target.value)} 
                            />
                    </div>   
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {pagesfiltered.map((page,idx)=>
                                <li key={idx}>
                                    <PageComponent page={page} siteurl={siteURL} sitePages={sitePages[0]}/>
                                </li>
                            )}
                        </ul>
                    </div>
            </div>
            }
            {tabSelected === 4 && 
            <div className={styles.detailsWrapper}>
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="pageName" 
                            placeholder="Filter by the library name"
                            onChange={e => searchLibFilter(e.target.value)} 
                            />
                    </div>
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {libfiltered.map((list,idx)=>
                                <li key={idx} style={{zIndex:'1'}}>
                                    <LibraryComponent list={list} siteurl={siteURL}/>
                                </li>
                            )}
                        </ul>
                    </div>
            </div>
            }
            {tabSelected === 5 && 
            <div className={styles.detailsWrapper}>
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="pageName" 
                            placeholder="Filter by the list name"
                            onChange={e => searchListFilter(e.target.value)} 
                            />
                    </div>
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {listfiltered.map((list,idx)=>
                                <li key={idx}>
                                    <ListComponent list={list} siteurl={siteURL}/>
                                </li>
                            )}
                        </ul>
                    </div>
            </div>
            }
            </>}
        </div>
    );
  }

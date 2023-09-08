import * as React from 'react';
import {useState, useEffect, useRef} from 'react';

//STYLING
import styles from './SiteOverviewMsf.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

//API
import { SPFx, graphfi} from "@pnp/graph";
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

import "@pnp/sp-admin";

import { Web } from "@pnp/sp/webs";   

import { SearchResults } from "@pnp/sp/search";
import { Site } from "@pnp/graph/sites";

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
        context
      } = props.details;

    const sp = props.sp//spfi().using(SPFxsp(context))
    //console.log(context)
//LOADERS      
    const [hubLoading,setHubLoading] = useState(true)
    const [subLoading,setSubLoading] = useState(true)
    const [listsLoading,setListsLoading] = useState(true)
    const [pagesLoading,setPagesLoading] = useState(true)

 
//CONST
    const [siteTitle,setSiteTitle] = useState(context.pageContext.web.title)
        
    const siteURL = site_url === undefined || site_url === null || site_url === "" ? context.pageContext.site.absoluteUrl : site_url

    const [siteID,setSiteID] = useState(
        site_id === undefined || site_id === null || site_id === "" ? context.pageContext.site.id._guid : site_id)

    const [tabSelected,setTabSelected] = useState(1)
    const tabSelectedHandler = (tab) => {
        setTabSelected(tab)
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

    async function getReport() {
        const response = await fetch(`${siteURL}/_api/site/usage`);
        const rawData = await response.text()
        
        return rawData
    }   
  
    async function getPages(id) {   
        setPagesLoading(true)
        const site = Web([sp.web, `${siteURL}`])      
        const sites = await site.lists.getByTitle("Site Pages").items.select('FileLeafRef', 'Title', 'Id', 'GUID')()
        setPagesLoading(false)
        
        return sites
    }   

    async function getLists() {
       
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
       
        getHub(siteID).then(result => {
            setHubsites([]);
            const arr:any = result
            setHubsites(arr);
        });

        getSubsites().then(result => {
            setSubsites([]);
            const arr:any = result
            setSubsites(arr);
        });

        getPages(siteID).then(result => {     
            setPages([]);
            const arr:any = result
            setPages(arr);
        })  

        /*
        getAdvancedReport(siteID).then(result => {
            result.map(feature => console.log(feature))
        }
        )
*/
        getReport().then(result => {
            const parser = new DOMParser()
            const xml = parser.parseFromString(result,"application/xhtml+xml")
            const defaultNamespaceURI = 'http://schemas.microsoft.com/ado/2007/08/dataservices';
            const elementName = 'Storage';
            const storageElement = xml.getElementsByTagNameNS(defaultNamespaceURI, elementName)[0];         
            const storage = Number(storageElement.textContent) / (1024 * 1024 * 1024)

            setStorage(storage)
        });

        getLists().then(result => {     
            setLists([]);
            const arr:any = result
            setLists(arr);
          })

       }, [site_id, site_url, siteID]);

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

//FILTERS    
    const genlist = lists.filter( list => list.template === 100 && list.name!=="DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES")
    const libraries = lists.filter( lib => lib.template === 101)
    const sitePages = lists.filter(lib => lib.template === 119 )   

    const [subFilter,setSubFilter] = useState("")
    const searchSubFilter = (e) => {
           setSubFilter(e)
       }
    const subsitesfiltered = subFilter === "" ? subsites : subsites.filter( sub => sub.Title.toLowerCase().includes(subFilter.toLowerCase()))

    const [hubFilter,setHubFilter] = useState("")
    const searchHubFilter = (e) => {
           setHubFilter(e)
       }    
    const hubsitesfiltered = hubFilter === "" ? hubsites : hubsites.filter( hub => hub.Title.toLowerCase().includes(hubFilter.toLowerCase()))

    const [pageFilter,setPageFilter] = useState("")
    const searchPageFilter = (e) => {
        setPageFilter(e)
    }
    //console.log(pages.length === 0)
    const pagesfiltered = pages.length === 0 ? [] : pageFilter === "" ? pages : pages.filter( page => page.Title.toLowerCase().includes(pageFilter.toLowerCase()))

    const [libFilter,setLibFilter] = useState("")
    const searchLibFilter = (e) => {
        setLibFilter(e)
    }
    const libfiltered = libFilter === "" ? libraries : libraries.filter( lib => lib.name.includes(libFilter))

    const [listFilter,setListFilter] = useState("")
    const searchListFilter = (e) => {
        setListFilter(e)
    }
    const listfiltered = listFilter === "" ? genlist : genlist.filter( list => list.name.includes(listFilter))

//MODALS   

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
 
    //console.log(pages)
    //console.log(pagesfiltered)
    //console.log(libfiltered)


     return (
        <div className={styles.overviewWrapper}>
            <div className={styles.mainSiteBox}>
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
                        <div className={styles.mainSiteBoxBottomLeft}>
                            <a className={styles.buttonModern} href={`${siteURL}/_layouts/15/viewlsts.aspx?view=14`} title="Site Content"><Icon iconName="AllApps"/></a>
                            <a className={styles.buttonClassic} href={`${siteURL}/_layouts/15/viewlsts.aspx`} title="Site Content (classic)"><Icon iconName="AllApps"/></a>  
                        </div>
                        <a href={`${siteURL}/_layouts/15/settings.aspx`} title="Site Settings"><Icon iconName="Settings"/></a>
                        <a href={`${siteURL}/_layouts/15/user.aspx`} title="Site Permissions"><Icon iconName="SecurityGroup"/></a>
                        <div className={styles.mainSiteBoxBottomLeft}>
                            <a className={styles.buttonModern} href={`${siteURL}/_layouts/15/appStore.aspx`} title="App store"><Icon iconName="Puzzle"/></a>
                            <a className={styles.buttonClassic} href={`${siteURL}/AppCatalog/Forms/AllItems.aspx`} title="Site App cataloge"><Icon iconName="Puzzle"/></a>
                        </div>
                        <a href={`${siteURL}/_layouts/15/siteanalytics.aspx?view=19`} title="Site Usage"><Icon iconName="LineChart"/></a>  
                        <a href={`${siteURL}/_layouts/15/storman.aspx`} title="Site Storage"><Icon iconName="OfflineStorage"/></a> 
                        <div className={styles.mainSiteBoxBottomLeft}>
                            <a className={styles.buttonModern} href={`${siteURL}/_layouts/15/AdminRecycleBin.aspx`} title="Site Recycle Bin"><Icon iconName="RecycleBin"/></a>
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
            {permVis && <PermissionsComponent onCloseHandler={permVisHandler} context={context} url={siteURL} sp={props.sp}/>}
            {pnpVis && <PnP_Generator onCloseHandler={pnpVisHandler} type={"top_site"} siteurl={site_url}/>}
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
                                <li key={idx}>
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

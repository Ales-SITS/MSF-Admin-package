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

import "@pnp/sp/items";
import "@pnp/sp/items/get-all";

import { Web } from "@pnp/sp/webs";   

import { SearchResults } from "@pnp/sp/search";
import { Site } from "@pnp/graph/sites";

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



//LOADERS      
    const [hubLoading,setHubLoading] = useState(true)
    const [subLoading,setSubLoading] = useState(true)
    const [listsLoading,setListsLoading] = useState(true)
    const [pagesLoading,setPagesLoading] = useState(true)

 
//CONST
    const [siteTitle,setSiteTitle] = useState(context.pageContext.web.title)
        
    const [siteURL,setSiteURL] = useState(
        site_url === undefined || site_url === null || site_url === "" ? context.pageContext.site.absoluteUrl : site_url)
    const [siteID,setSiteID] = useState(
        site_id === undefined || site_id === null || site_id === "" ? context.pageContext.site.id._guid : site_id)

//CONST & HIDERS        
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

    const [pages, setPages] = useState([])
    const [pageshide, setPageshide] = useState(expanded)
    const pageshideHandler = () => {
        setPageshide(!pageshide)
    }

    const [pagesfiltered, setPagesfiltered] = useState([])
    //const pagesfiltered = pages.length === 0 ? [] : pages.filter( page => page.Title.includes(pageFilter))


    const [lists,setLists] = useState([])
    const [libhide,setLibhide] = useState(expanded)
    const libhideHandler = () => {
        setLibhide(!libhide)
    }

    const [lishide,setLishide] = useState(expanded)
    const lishideHandler = () => {
        setLishide(!lishide)
    }

//GETTERS

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
  
    async function getPages(id) {   
        setPagesLoading(true)
        const sp = spfi().using(SPFxsp(context));     
        const site = Web([sp.web, `${siteURL}`])      
        const sites = await site.lists.getByTitle("Site Pages").items.select('FileLeafRef', 'Title', 'Id', 'GUID')()
        setPagesLoading(false)
        sites.length === 0 ? setPagesfiltered(pages) :  setPagesfiltered(pages.filter( page => page.Title.includes(pageFilter)))

        return sites
    }   

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

        getPages(site_id).then(result => {     
            setPages([]);
            const arr:any = result
            setPages(arr);
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

//FILTERS    
    const genlist = lists.filter( list => list.template === 100 && list.name!=="DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CONTENTTYPES")
    const libraries = lists.filter( lib => lib.template === 101)
    const sitePages = lists.filter(lib => lib.template === 119 )   

    const [subFilter,setSubFilter] = useState("")
    const searchSubFilter = (e) => {
           setSubFilter(e)
       }
    const subsitesfiltered = subsites.filter( sub => sub.Title.includes(subFilter))

    const [hubFilter,setHubFilter] = useState("")
    const searchHubFilter = (e) => {
           setHubFilter(e)
       }    
    const hubsitesfiltered = hubsites.filter( hub => hub.Title.includes(hubFilter))

    const [pageFilter,setPageFilter] = useState("")
    const searchPageFilter = (e) => {
        setPageFilter(e)
    }
    //console.log(pages.length === 0)
    //const pagesfiltered = pages.length === 0 ? [] : pages.filter( page => page.Title.includes(pageFilter))

    const [libFilter,setLibFilter] = useState("")
    const searchLibFilter = (e) => {
        setLibFilter(e)
    }
    const libfiltered = libraries.filter( lib => lib.name.includes(libFilter))

    const [listFilter,setListFilter] = useState("")
    const searchListFilter = (e) => {
        setListFilter(e)
    }
    const listfiltered = genlist.filter( list => list.name.includes(listFilter))

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
                    <span>{hubhide ? "▶ " : "▼ "} Hub associated sites</span>
                    {
                    hubLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span><span className={styles.displayedNum}>{hubsitesfiltered.length}/</span>{hubsites.length}</span>
                     }
                </button>
                {!hubhide && 
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="siteName" 
                            placeholder="Filter by the site title"
                            onChange={e => searchHubFilter(e.target.value)} 
                            />
                    </div>
                    }
                {!hubhide &&
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {hubsitesfiltered.map((hubsite,idx)=>
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
                    <span>{subhide ? "▶ " : "▼ "} Subsites</span>
                    {
                    subLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span><span className={styles.displayedNum}>{subsitesfiltered.length}/</span>{subsites.length}</span>
                    }
                    
                </button>
                {!subhide && 
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="subsiteName" 
                            placeholder="Filter by the subsite title"
                            onChange={e => searchSubFilter(e.target.value)} 
                            />
                    </div>
                    }
                {!subhide && 
                    <div className={styles.resultsWrapper}>

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
                <button className={styles.detailsWrapperButton} onClick={pageshideHandler}>
                    <span>{lishide ? "▶ " : "▼ "} Pages</span>
                    {
                    pagesLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span><span className={styles.displayedNum}>{pagesfiltered.length}/</span>{pages.length}</span>
                    }
                </button>
                {!pageshide && sitePages[0] !== undefined && 
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="pageName" 
                            placeholder="Filter by the page title"
                            onChange={e => searchPageFilter(e.target.value)} 
                            />
                    </div>
                    }
                {!pageshide && sitePages[0] !== undefined &&
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {pagesfiltered.map((page,idx)=>
                                <li key={idx}>
                                    <PageComponent page={page} siteurl={siteURL} sitePages={sitePages[0]}/>
                                </li>
                            )}
                        </ul>
                    </div>
                }
            </div>
            <div className={styles.detailsWrapper}>
                <button className={styles.detailsWrapperButton} onClick={libhideHandler}>
                    <span>{libhide ? "▶ " : "▼ "} Libraries</span>
                    {
                    listsLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span><span className={styles.displayedNum}>{libfiltered.length}/</span>{libraries.length}</span>
                    }
                </button>
                {!libhide && 
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="pageName" 
                            placeholder="Filter by the library name"
                            onChange={e => searchLibFilter(e.target.value)} 
                            />
                    </div>
                    }
                {!libhide && 
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {libfiltered.map((list,idx)=>
                                <li key={idx}>
                                    <LibraryComponent list={list} siteurl={siteURL}/>
                                </li>
                            )}
                        </ul>
                    </div>
                }
            </div>
            <div className={styles.detailsWrapper}>
                <button className={styles.detailsWrapperButton} onClick={lishideHandler}>
                    <span>{lishide ? "▶ " : "▼ "} Lists</span>
                    {
                    listsLoading ? 
                    <div className={styles.loader}><div></div><div></div><div></div><div></div></div>:
                    <span><span className={styles.displayedNum}>{listfiltered.length}/</span>{genlist.length}</span>
                    }
                </button>
                {!lishide && 
                    <div className={styles.resultsFilterInputBox}>
                        <input
                            className={styles.resultsFilterInput} 
                            type="text" 
                            name="pageName" 
                            placeholder="Filter by the list name"
                            onChange={e => searchListFilter(e.target.value)} 
                            />
                    </div>
                    }
                {!lishide && 
                    <div className={styles.resultsWrapper}>
                        <ul>
                            {listfiltered.map((list,idx)=>
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

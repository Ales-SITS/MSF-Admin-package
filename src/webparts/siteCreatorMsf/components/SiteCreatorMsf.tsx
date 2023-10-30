import *  as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';

import M365 from './M365'
import NonM365 from './NonM365'
import Communication from './Communication'
import Declined from './Declined'

//API
import { SPFx,graphfi } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";

import { SPFx as SPFxsp, spfi } from "@pnp/sp";
import "@pnp/sp/site-designs";
import "@pnp/sp/hubsites";


export default function SiteCreatorMsf (props) {

  const graph = graphfi().using(SPFx(props.context))
  const sp = spfi().using(SPFxsp(props.context))
  const siteurl = props.context.pageContext.site.absoluteUrl

  const domains = {
    "aerzte-ohne-grenzen.at":"VIE",
    "amman.msf.org":"UNCLEAR",
    "amsterdam.msf.org":"OCA",
    "athens.msf.org":"ATH",
    "azg.be":"OCB",
    "barcelona.msf.org":"OCBA",
    "berlin.msf.org":"BER",
    "bi.msf.org":"Shared",
    "beijing.msg.org":"HKG",
    "bogota.msf.org":"OCBA",
    "bordeaux.msf.org":"BDX",
    "brussels.msf.org":"OCB",
    "buenosaires.msf.org":"OCBA",
    "copenhagen.msf.org":"CPH",
    "dubai.msf.org":"OCP",
    "dublin.msf.org":"LDN",
    "epicentre.msf.org":"EPI",
    "geneva.msf.org":"OCG",
    "zurich.msg.org":"OCG",
    "helsinki.msf.org":"HEL",
    "hongkong.msf.org":"HKG",
    "china.msf.org":"HKG",
    "istanbul.msf.org":"OCBA",
    "joburg.msf.org":"JNB",
    "lakareutangranser.se":"STO",
    "legerutengrenser.no":"OSL",
    "legerutengrenser.teams.telia.no":"OSL",
    "london.msf.org":"LDN",
    "luxembourg.msf.org":"LUX",
    "madrid.msf.org":"OCBA",
    "mail.msf.org":"Shared",
    "manila.msf.org":"HKG",
    "mexico.msf.org":"OCG",
    "migration.msf.org":"TOR",
    "mmsui.in":"OCB",
    "montevideo.msf.org":"OCBA",
    "moscow.msf.org":"BER",
    "msf.be":"OCB",
    "msf.ca":"TOR",
    "msf.dk":"CPH",
    "msf.global":"Shared",
    "msf.no":"OSL",
    "msf.org":"Shared",
    "msfintl.onmicrosoft.com":"Shared",
    "nairobi.msf.org":"OCBA",
    "new-delhi.msf.org":"OCA",
    "newyork.msf.org":"NYC",
    "nz.msf.org":"SYD",
    "oca.msf.org":"OCA",
    "oslo.msf.org":"OSL",
    "paris.msf.org":"OCP",
    "philippines.msf.org":"HKG",
    "prague.msf.org":"PRG",
    "rio.msf.org":"RIO",
    "rome.msf.org":"ROM",
    "seasia.msf.org":"HKG",
    "seeap.msf.org":"HKG",
    "seoul.msf.org":"SEL",
    "singapore.msf.org":"HKG",
    "sits.msf.org":"SITS",
    "somalia.msf.org":"OCA",
    "somaliland.msf.org":"OCA",
    "stockholm.msf.org":"STO",
    "sydney.msf.org":"SYD",
    "taipei.msf.org1":"TPE",
    "taiwan.msf.org":"TPE",
    "tokyo.msf.org":"TYO",
    "toronto.msf.org":"TOR",
    "tunis.msf.org":"UNCLEAR",
    "vienna.msf.org":"VIE",
    "waca.msf.org":"WaCA"
  }

  const domain = `${domains[`${props.context.pageContext.user.email?.split("@")[1].toLowerCase()}`]}`

  const [selectedType,setSelectedType] = useState(1)
  const selectedTypeHandler = (type) => {
    setSelectedType(type)
  }

  const [isAdmin, setIsAdmin] = useState(false)
  const [loader,setLoader] = useState(true)
  const [contenttypes,setContenttypes] = useState([])
  const [sitedesigns,setSitedesigns] = useState([])
  const [hubsites,setHubsites] = useState([])

  useEffect(()=>{
    getAdminGroup()
  },[])

  const getAdminGroup = async() => {
    const admingroups = await graph.me.memberOf();
    const allowgroup = await graph.groups.getById("d35fc400-84f4-40c1-956f-c0532a976d1f").members();
    for (const grpA of admingroups) {
      if (allowgroup.some(grpB => grpA.id === grpB.id)) {
        setIsAdmin(true)
      }   
    }
    setLoader(false)
  }

  useEffect(()=>{
    loadContentTypes()
    loadSiteDesign()
    loadHubs()
  },[loader])
 
  async function loadContentTypes() {
    const urlObject = new URL(siteurl);
    const host = urlObject.hostname
    const path = urlObject.pathname
    const hubSiteContentTypes = await graph.sites.getByUrl(host, path).contentTypes.getCompatibleHubContentTypes();
    const cleanedCT = hubSiteContentTypes.map(({id:key, name:text})=>({
      key,
      text
    }))
    setContenttypes(cleanedCT)
  }

  async function loadSiteDesign(){
    const allSiteDesigns = await sp.siteDesigns.getSiteDesigns();
    const cleanedSD = allSiteDesigns.map(({ Id:key, Title:text, WebTemplate }) => ({
      key,
      text,
      WebTemplate
    }))

    setSitedesigns(cleanedSD)
  }

  async function loadHubs() {
    const hubsites = await sp.hubSites();
    const cleanedhubs = hubsites.filter(hub => hub.Targets!==null).map(({ID:key, Title:text})=>({
      key,
      text  
    }))
    setHubsites(cleanedhubs)
  }


  return (
    <div className={styles.app_wrapper}>
      <h2>Site Creator</h2>
      <div className={styles.app_navigation}>
        <button onClick={()=>selectedTypeHandler(1)} className={selectedType === 1 && styles.selected_site_type}>Teams site (M365)</button>
        <button onClick={()=>selectedTypeHandler(2)} className={selectedType === 2 && styles.selected_site_type}>Team site without M365</button>
        <button onClick={()=>selectedTypeHandler(3)} className={selectedType === 3 && styles.selected_site_type}>Communication site</button>
      </div>
      {!isAdmin ? <Declined context={props.context} loader={loader}/> :
       selectedType === 1 ?
        <M365 
        context={props.context} 
        ct={contenttypes}  
        hs={hubsites} 
        sd={sitedesigns.filter((sd)=>sd.WebTemplate === "64")}
        domain={domain}
        /> :
       selectedType === 2 ?
       <NonM365 
       context={props.context} 
       ct={contenttypes}  
       hs={hubsites} 
       sd={sitedesigns.filter((sd)=>sd.WebTemplate === "1")}
       domain={domain}
       /> : 
       <Communication 
       context={props.context} 
       ct={contenttypes} 
       hs={hubsites} 
       sd={sitedesigns.filter((sd)=>sd.WebTemplate === "68")}
       domain={domain}
       />
      }
    </div>
  )
}


/*
Id                  : f5ce4b3c-7b29-44e5-9a8d-cdd8ad2db50b
Title               : External sharing ExistingExternalUserSharingOnly
Description         : Sets External sharing to ExistingExternalUserSharingOnly 
Content             : 
Version             : 1
IsSiteScriptPackage : False

Id                  : 721f126f-a657-4f38-8e44-4ddca33bb8be
Title               : External sharing Disabled
Description         : Sets External sharing to Disabled
Content             : 
Version             : 1
IsSiteScriptPackage : False

Id                  : 6563274d-f5fe-451d-a916-f91e488c86eb
Title               : External sharing ExternalUserSharingOnly
Description         : Sets External sharing to ExternalUserSharingOnly
Content             : 
Version             : 1
IsSiteScriptPackage : False

Id                  : 3897ba25-22bd-40ad-9fb3-a2df5132c928
Title               : External sharing ExternalUserAndGuestSharing
Description         : Sets External sharing to ExternalUserAndGuestSharing
Content             : 
Version             : 1
IsSiteScriptPackage : False

*/
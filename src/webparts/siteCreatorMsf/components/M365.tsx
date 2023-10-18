import * as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';
import { ISiteCreatorMsfProps } from './ISiteCreatorMsfProps';
import {  PeoplePicker } from '@microsoft/mgt-react';

//API
import { MSGraphClientV3  } from "@microsoft/sp-http";

//PNP SP
import { spfi, SPFx as SPFxsp} from "@pnp/sp";
import "@pnp/sp/site-designs"
import "@pnp/sp/sites";

import { IHubSiteInfo } from  "@pnp/sp/hubsites";
import "@pnp/sp/hubsites";

//PNP GRAPH
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";


export default function M365 (props) {

  //API
  const context = props.context
  const sp = spfi().using(SPFxsp(context))
  const graph = graphfi().using(SPFx(context))

  //General
  const domain = context.pageContext.user.email?.split("@")[1]?.split(".")[0]?.toUpperCase()

  const [progress, setProgress] = useState("Not run yet")
  const [error, setError] = useState ("")
  const [title,setTitle] = useState("")
  const [titleExist, setTitleExist] = useState(false)
  const addTitle = (e) => {
    setTitle(e.target.value)
    siteExistsChecker(e.target.value)
  }


  //Users states
  const[adminId,setAdminId] =useState([])
  const [owners,setOwners] = useState([])
  const [members,setMembers] = useState([])

  const addOwners = (e) => {
    const ownersMails = e.detail.map(owner => `https://graph.microsoft.com/v1.0/users/${owner.id}`)
    setOwners([...ownersMails])
  }

  const addMembers = (e) => {
    const membersMails = e.detail.map(member => `https://graph.microsoft.com/v1.0/users/${member.id}`)
    setMembers(membersMails)
  }

  //Admin check
  useEffect(()=>{
    getAdmin().then(result => {
      setAdminId([result.id])
  })
  },
  [])

  //Hub states
  const [hub,setHub] = useState("")
  const [hubTitle,setHubTitle] = useState("")
  const [hubOwner, setHubOwner] = useState(false)
  const [hubChecker,setHubChecker] = useState(false)

  const addHub = (e) => {
    setHub(e.target.value)
  }
  
  //Hub checkers
  useEffect(()=> {
    if (hub === "") {
      setHubChecker(true)
    } else {
    const hubcheck = hub.split("-")
    hubcheck.length != 5 || 
    hubcheck[0].length !=8 ||
    hubcheck[1].length !=4 ||
    hubcheck[2].length !=4 ||
    hubcheck[3].length !=4 ||
    hubcheck[4].length !=12 ? 
    setHubChecker(false)  : setHubChecker(true)
    }
  },[hub])

  useEffect(()=>{
    hubChecker && getHub(hub)
  },[hubChecker,hub])

  async function getHub (hubID) {
    const hubsite: IHubSiteInfo = await sp.hubSites.getById(hubID)();
    setHubTitle(hubsite.Title)
    
    hubsite.Targets.includes(`${context.pageContext.user.email.toLowerCase()}`) ? setHubOwner(true) : setHubOwner(false)
  }

  //Design states
  const [siteDesign, setSiteDesign] = useState("")
  const [siteDesignTitle, setSiteDesignTitle] = useState("")
  const [designChecker,setDesignChecker] = useState(false)

  const addDesign = (e) => {
    setDesignChecker(false)
    setSiteDesign(e.target.value)
  }

  //Design checker
  useEffect(()=> {
    if (siteDesign === "") {
      setDesignChecker(true)
    } else {
    const designcheck = siteDesign.split("-")
    designcheck.length != 5 || 
    designcheck[0].length !=8 ||
    designcheck[1].length !=4 ||
    designcheck[2].length !=4 ||
    designcheck[3].length !=4 ||
    designcheck[4].length !=12 ? 
    setDesignChecker(false)  : setDesignChecker(true)
  }
  },[siteDesign])

  useEffect(()=>{
    designChecker && siteDesign === "" ? null : designChecker && siteDesign !== "" ? getDesign(siteDesign) : null
  },[designChecker, siteDesign])

  async function getDesign (designID) {
       const design = await sp.siteDesigns.getSiteDesignMetadata(designID)
      setSiteDesignTitle(design.Title)
      console.log("Design found")
   }

  //Privacy state
  const [privacy, setPrivacy] = useState("Private")

  const onPrivacyChange = e => {
    setPrivacy(e.target.value)
  }

  //Sharing state
  const [sharingId, setSharingId] = useState("123ac0ed-b076-4507-82e4-de444923a4b5")
  const [sharing, setSharing] = useState("New and existing guest")
  const onSharingChange = e => {
    setSharing(e.target.value)
    e.target.value === "Anyone" ?
    setSharingId("004bbf35-ed4a-45e6-9199-cc9881aeba64") :
    e.target.value === "New and existing guest" ?
    setSharingId("123ac0ed-b076-4507-82e4-de444923a4b5") :
    e.target.value === "Existing guest only" ?
    setSharingId("513beb79-9027-4297-a38d-db7cbfb83b07") :
    setSharingId("3b0ffe11-9840-4561-96a9-c6c417976db9")
  }

// SCRIPT
/*
Id                  : 3b0ffe11-9840-4561-96a9-c6c417976db9 OK
Title               :  External sharing Disabled (Team Site)
WebTemplate         : 64
SiteScriptIds       : {721f126f-a657-4f38-8e44-4ddca33bb8be}
Description         : Sets External sharing to Disabled (Team Site)

Id                  : 6bf7f2ba-2f3f-4449-9c7d-596f22e9cb83
Title               :  External sharing ExistingExternalUserSharingOnly (Team Site)
WebTemplate         : 64
SiteScriptIds       : {f5ce4b3c-7b29-44e5-9a8d-cdd8ad2db50b}
Description         : Sets External sharing to ExistingExternalUserSharingOnly (Team Site)

Id                  : 513beb79-9027-4297-a38d-db7cbfb83b07 OK
Title               :  External sharing ExternalUserSharingOnly (Team Site)
WebTemplate         : 64
SiteScriptIds       : {6563274d-f5fe-451d-a916-f91e488c86eb}
Description         : Sets External sharing to ExternalUserSharingOnly (Team Site)

Id                  : 004bbf35-ed4a-45e6-9199-cc9881aeba64 OK
Title               :  External sharing ExternalUserAndGuestSharing (Team Site)
WebTemplate         : 64
SiteScriptIds       : {3897ba25-22bd-40ad-9fb3-a2df5132c928}
Description         : Sets External sharing to ExternalUserSharingOnly (Team Site)
*/

//SITE CREATION
  const createSite = (e) => {
    e.preventDefault()
     const group = {
      description: 'For API testing purposes',
      displayName: `GRP-${domain}-${title}`,
      groupTypes: ['Unified'],
      mailEnabled: true,
      securityEnabled: false,
      visibility: `${privacy}`,
      mailNickname: `GRP-${domain}-${title}`
    };

    if (members.length > 0) {
      group['members@odata.bind'] = [...members];
    }

    if (owners.length > 0) {
      group['owners@odata.bind'] = [...owners];
    }

    // Log the current operation
    setProgress("Creating Team site (M365) ...");
  
    context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        client.api('/groups')
	      .post(group) 
        .then(() => { 
          siteExists(`https://msfintl.sharepoint.com/sites/GRP-${domain}-${title}`)      
        })
      .catch((error: any) => {
          setError(`Error when creating site: ${error}`);
      });
      })    
  };

  const siteExistsChecker = async(titlecheck) => {
      try {
        const exists = await sp.site.exists(`https://msfintl.sharepoint.com/sites/GRP-${domain}-${titlecheck}`)
        setTitleExist(exists)
      } catch (error) {
        console.error('Site exists:', error);
      }
  }

  const siteExists = async(siteUrl) => {
      let exists = false

      while(!exists) {
        exists =  await sp.site.exists(siteUrl)
        if(!exists){
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }
      
      setProgress("Team site (M365) created. Preparing other settings ...");
      await new Promise((resolve) => setTimeout(resolve, 10000));

      await applySharing(siteUrl);
      hub !== "" && hubChecker ? await associateToHub(siteUrl) : null
      siteDesign !== "" && designChecker ? await applyScript(siteUrl) : null
  }

  async function applySharing(siteUrl) {
    setProgress("Applying sharing settings ...");
    try {
        await sp.siteDesigns.applySiteDesign(
          `${sharingId}`,
          `${siteUrl}`
        );
        setProgress("Other settings and scripts applied")
      } catch (error) {
        setError(`Error when applying sharing settings: ${error}`);
      }
  } 

  async function applyScript(siteUrl) {
        setProgress("Applying site design ...");
        try {
            await sp.siteDesigns.applySiteDesign(
              `${siteDesign}`,
              `${siteUrl}`
            );
            setProgress("Other settings and scripts applied")
          } catch (error) {
            setError(`Error when applying site design: ${error}`);
          }
  } 

  async function associateToHub (siteUrl) {
      setProgress("Associating with hub ...");
        const newsp = spfi(siteUrl).using(SPFxsp(context))
        try {
          await newsp.site.joinHubSite(`${hub}`)
          setProgress("Finished")
        } catch (error) {
          setError(`Error when associating to hub: ${error}`);
        }
  }

  async function getAdmin(){
    
    const admin = await graph.me()

    return admin
  }

  const disabled = title === "" || titleExist || !hubChecker || !designChecker ? true : false

  return (
    <div className={styles.site_wrapper}>
      <form className={styles.form_wrapper} onSubmit={createSite}>
          <label htmlFor='siteTitle'>Site name</label>
          <input id="siteTitle" type="text" onChange={addTitle}/>
          <span className={styles.input_comment}>{title === "" ? "Type a site title" : titleExist ? `NG, GRP-${domain}-${title} already exists` : "OK"}</span>
          <span>M365 group Owners</span>
          <PeoplePicker defaultSelectedUserIds={adminId} selectionMode="multiple" selectionChanged={addOwners}/>
          <span>M365 group Members</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addMembers}/>
          <div className={styles.selection_box}>
            <h4>Privacy</h4>
            <span>
                <input
                    type="radio"
                    name="privacy"
                    value= "Private"
                    id="Private"
                    checked={privacy === "Private"}
                    onChange={onPrivacyChange}
                />
                <label htmlFor="Private">Private</label>
            </span>
            <span>
                <input
                    type="radio"
                    name="privacy"
                    value= "Public"
                    id="Public"
                    checked={privacy === "Public"}
                    onChange={onPrivacyChange}
                />
                <label htmlFor="Public">Public</label>
            </span>
          </div>
          <div className={styles.selection_box}>
            <h4>External sharing <a target="_blank" href="https://learn.microsoft.com/en-US/sharepoint/change-external-sharing-site?WT.mc_id=365AdminCSH_inproduct#which-option-to-select">?</a></h4>
              <span>
                <input
                  type="radio"
                  name="sharing"
                  value= "Anyone"
                  id="Anyone"
                  checked={sharing === "Anyone"}
                  onChange={onSharingChange}
                />
                <label htmlFor="Anyone">Anyone</label>
              </span>
              <span>
                <input
                    type="radio"
                    name="sharing"
                    value= "New and existing guest"
                    id="New and existing guest"
                    checked={sharing === "New and existing guest"}
                    onChange={onSharingChange}
                />
                <label htmlFor="New and existing guest">New and existing guest</label>
              </span>
              <span>
                <input
                    type="radio"
                    name="sharing"
                    value= "Existing guest only"
                    id="Existing guest only"
                    checked={sharing === "Existing guest only"}
                    onChange={onSharingChange}
                />
                <label htmlFor="Existing guest only">Existing guest only</label>
              </span>
              <span>
                <input
                    type="radio"
                    name="sharing"
                    value= "Only people in your organization"
                    id="Only people in your organization"
                    checked={sharing === "Only people in your organization"}
                    onChange={onSharingChange}
                />
                <label htmlFor="Only people in your organization">Only people in your organization</label>
              </span>
          </div>
          <label htmlFor='siteScript'>Custom site design (design id) <a target="_blank" href="https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview">?</a></label>
          <input id="siteScript" type="text" placeholder='00000000-0000-0000-0000-000000000000' onChange={addDesign}/>
          <span className={styles.input_comment}>
            { 
            siteDesign === "" ? "No design applied - OK" :
            designChecker ? 
              `${siteDesignTitle} - OK` : 
              "Wrong format or id"
            }
          </span>
          <label htmlFor='hubId'>Associate with hub (site id)</label>
          <input id="hubId" type="text" placeholder='00000000-0000-0000-0000-000000000000' onChange={addHub}/>
          <span className={styles.input_comment}>
            { 
            hub === "" ? "Not associated to any hub - OK" :
            hubChecker ? 
              `${hubTitle} ${hubOwner ? "- OK" : "- Your account is not an owner of the hub or cannot associate sites to it!"}` : 
              "Wrong format or id"
            }
          </span>
          <div className={styles.createSite_button_wrapper}>
            <input className={styles.createSite_button} type="submit" onClick={createSite} value="Create site" 
                   disabled = {disabled}/>
          </div>
      </form>
      <div className={styles.result_wrapper}>
        <div className={styles.result_list}>
          <p>You will create a site with M365 group, which includes planner, teams etc. Your site will have the following properties:</p>
          <h3>GRP-{domain}-{title}</h3>
          <span>Url: https://msfintl.sharepoint.com/sites/GRP-{domain}-{title}</span>
          <span>Privacy: {privacy}</span>
          <span>Sharing: {sharing}</span>
          <span>Site design: {siteDesignTitle}</span>
          <span>Associate with hub: {hubTitle}</span>
        </div>
        <div className={styles.result_progress}>
          {progress === "Finished" ? 
          <a target="_blank" href={`https://msfintl.sharepoint.com/sites/GRP-${domain}-${title}`}>Finished - click to open</a> :
          <span>{progress}</span>
          }
          {error}      
        </div>
      </div>
    </div>
  )
}

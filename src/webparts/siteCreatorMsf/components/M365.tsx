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

  const context = props.context
  const sp = spfi().using(SPFxsp(context))
  const graph = graphfi().using(SPFx(context))

  const domain = context.pageContext.user.email?.split("@")[1]?.split(".")[0]?.toUpperCase()

  const [title,setTitle] = useState("")
  const [titleExist, setTitleExist] = useState(false)
  const addTitle = (e) => {
    setTitle(e.target.value)
    siteExistsChecker(e.target.value)
  }

  const [siteDesign, setSiteDesign] = useState("")
  const addDesign = (e) => {
    setSiteDesign(e.target.value)
  }

  const[adminId,setAdminId] =useState([])
  const [owners,setOwners] = useState([])
  const [members,setMembers] = useState([])

  const [hub,setHub] = useState("")
  const addHub = (e) => {
    setHub(e.target.value)
  }
  
  const [hubTitle,setHubTitle] = useState("")
  const [hubOwner, setHubOwner] = useState(false)

  const [hubChecker,setHubChecker] = useState(false)

  const [designChecker,setDesignChecker] = useState(false)

  const [siteCreated,setSiteCreated] = useState(false)

  const [progress, setProgress] = useState("Not run yet")

  const addOwners = (e) => {
    const ownersMails = e.detail.map(owner => `https://graph.microsoft.com/v1.0/users/${owner.id}`)
    setOwners([...ownersMails])
  }

  const addMembers = (e) => {
    const membersMails = e.detail.map(member => `https://graph.microsoft.com/v1.0/users/${member.id}`)
    setMembers(membersMails)
  }

  const [privacy, setPrivacy] = useState("Private")

  const onPrivacyChange = e => {
    setPrivacy(e.target.value)
  }

  const [sharing, setSharing] = useState("New and existing guest")
  const onSharingChange = e => {
    setSharing(e.target.value)
  }

// Site script externalsharing disabled 
// de32b387-f7af-4d9f-8e82-40ad7fa23500

// Site design (team site)
// 123ac0ed-b076-4507-82e4-de444923a4b5


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
    setProgress("Creating basic site...");
  
    context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        client.api('/groups')
	      .post(group) 
        .then(() => { 
          siteExists(`https://msfintl.sharepoint.com/sites/GRP-${domain}-${title}`)      
        })
      .catch((error: any) => {
          console.error('POST error:', error);
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
      setSiteCreated(exists)
      setProgress("Basic site created, preparing other settings ...");
      await new Promise((resolve) => setTimeout(resolve, 10000));

      await applyScript(siteUrl);
      hub !== "" && hubChecker ? await associateToHub(siteUrl) : null
  }

  async function applyScript(siteUrl) {
        setProgress("Applying settings and scripts...");
        try {
            await sp.siteDesigns.applySiteDesign(
              "123ac0ed-b076-4507-82e4-de444923a4b5",
              `${siteUrl}`
            );
            setProgress("Other settings and scripts applied")
          } catch (error) {
            console.error('Error applying site design:', error);
          }
  } 

  async function associateToHub (siteUrl) {
      setProgress("Associating with hub ...");
        const newsp = spfi(siteUrl).using(SPFxsp(context))
        try {
          await newsp.site.joinHubSite(`${hub}`)
          setProgress("Finished")
        } catch (error) {
          console.error('Error associating to hub:', error);
        }
  }

  async function getAdmin(){
    
    const admin = await graph.me()

    return admin
  }

//Admin check
  useEffect(()=>{
    getAdmin().then(result => {
      setAdminId([result.id])
   })
  },
  [])

//hub checkers
  useEffect(()=> {
    const hubcheck = hub.split("-")
    hubcheck.length != 5 || 
    hubcheck[0].length !=8 ||
    hubcheck[1].length !=4 ||
    hubcheck[2].length !=4 ||
    hubcheck[3].length !=4 ||
    hubcheck[4].length !=12 ? 
    setHubChecker(false)  : setHubChecker(true)
  },[hub])

  useEffect(()=>{
    hubChecker && getHub(hub)
  },[hubChecker])
 
  async function getHub (hubID) {
    const hubsite: IHubSiteInfo = await sp.hubSites.getById(hubID)();
    setHubTitle(hubsite.Title)
    hubsite.Targets.includes(`${context.pageContext.user.email.toLowerCase()}`) ? setHubOwner(true) : setHubOwner(false)
  }

//design checker
  useEffect(()=> {
    const designcheck = hub.split("-")
    designcheck.length != 5 || 
    designcheck[0].length !=8 ||
    designcheck[1].length !=4 ||
    designcheck[2].length !=4 ||
    designcheck[3].length !=4 ||
    designcheck[4].length !=12 ? 
    setDesignChecker(false)  : setDesignChecker(true)
  },[siteDesign])

  useEffect(()=>{
    designChecker && getDesign(siteDesign)
  },[designChecker])

  async function getDesign (designID) {
    const design = await sp.siteDesigns.getSiteDesignMetadata(designID)
    console.log(design)
  }

  return (
    <div className={styles.m365_wrapper}>
      <form className={styles.form_wrapper} onSubmit={createSite}>
          <label htmlFor='siteTitle'>Site name</label>
          <input id="siteTitle" type="text" onChange={addTitle}/>
          <span>{titleExist ? "A site with this title already exists" : "OK"}</span>
          <span>Owners</span>
          <PeoplePicker defaultSelectedUserIds={adminId} selectionMode="multiple" selectionChanged={addOwners}/>
          <span>Members</span>
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
          <label htmlFor='siteScript'>Custom site design (id) <a target="_blank" href="https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview">?</a></label>
          <input id="siteScript" type="text" onChange={addDesign}/>
          <label htmlFor='hubId'>Associate with hub (site id)</label>
          <input id="hubId" type="text" placeholder='00000000-0000-0000-0000-000000000000' onChange={addHub}/>
          <span>
            { 
            hub === "" ? "" :
            hubChecker ? 
              `${hubTitle} ${hubOwner ? "- OK" : "- Your account is not an owner of the hub or cannot associate sites to it!"}` : 
              "Wrong format"
            }
          </span>
          <div className={styles.createSite_button_wrapper}>
            <input className={styles.createSite_button} type="submit" onClick={createSite} value="Create site" 
                   disabled = {titleExist || !hubChecker  ? true : false}/>
          </div>
      </form>
      <div className={styles.result_wrapper}>
        <div className={styles.result_list}>
          <p>You will create a site with M365 group, which includes planner, teams etc. Your site will have the following properties:</p>
          <h3>GRP-{domain}-{title}</h3>
          <span>Url: https://msfintl.sharepoint.com/sites/GRP-{domain}-{title}</span>
          <span>Privacy: {privacy}</span>
          <span>Sharing: {sharing}</span>
          <span>Associate with hub: {hubTitle}</span>
        </div>
        <div className={styles.result_progress}>
          {progress === "Finished" ? 
          <a target="_blank" href={`https://msfintl.sharepoint.com/sites/GRP-${domain}-${title}`}>Finished - click to open</a> :
          <span>{progress}</span>
          }
        </div>
      </div>
    </div>
  )
}

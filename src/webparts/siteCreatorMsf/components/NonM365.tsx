import * as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';
import {  PeoplePicker } from '@microsoft/mgt-react';

//PNP SP
import { spfi, SPFx as SPFxsp} from "@pnp/sp";
import { Web } from "@pnp/sp/webs";  
import { IHubSiteInfo } from  "@pnp/sp/hubsites";
import "@pnp/sp/hubsites";
import "@pnp/sp/sites";
import "@pnp/sp/user-custom-actions";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-designs";
import "@pnp/sp/features";

//PNP GRAPH
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";



export default function NonM365 (props) {

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
  const [visitors,setVisitors] = useState([])

  const addOwners = (e) => {
    const owners = e.detail.map(owner =>  `i:0#.f|membership|${owner.userPrincipalName.toLowerCase()}`)
    setOwners([...owners])
  }

  const addMembers = (e) => {
    const members = e.detail.map(member => `i:0#.f|membership|${member.userPrincipalName.toLowerCase()}`)
    setMembers(members)
  }

  const addVisitors = (e) => {
    const visitors = e.detail.map(member => `i:0#.f|membership|${member.userPrincipalName.toLowerCase()}`)
    setVisitors(visitors)
  }


  //Admin check
  useEffect(()=>{
    getAdmin().then(result => {
      setAdminId([result.id])
      setOwners([result.mail.toLowerCase()])
  })
  },
  [])

  async function getAdmin(){
      const admin = await graph.me()
    return admin
  }

  //Hub states
  const [hub,setHub] = useState("")
  const [hubTitle,setHubTitle] = useState("None")
  const [hubOwner, setHubOwner] = useState(false)
  const [hubChecker,setHubChecker] = useState(false)

  const addHub = (e) => {
    setHub(e.target.value)
  }
  
  //Hub checkers
  useEffect(()=> {
    if (hub === "") {
      setHubChecker(true)
      setHubTitle("None")
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
    hubChecker && hub === "" ? setHubTitle("None") :
    hubChecker && hub !== "" ? getHub(hub) :
    setHubTitle("None")
  },[hubChecker,hub])

  async function getHub (hubID) {
    const hubsite: IHubSiteInfo = await sp.hubSites.getById(hubID)();
    setHubTitle(hubsite.Title)
    
    hubsite.Targets.includes(`${context.pageContext.user.email.toLowerCase()}`) ? setHubOwner(true) : setHubOwner(false)
  }

  //Design states
  const [siteDesign, setSiteDesign] = useState("")
  const [siteDesignTitle, setSiteDesignTitle] = useState("None")
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

  //Sharing state
  const [sharingId, setSharingId] = useState("71d64c0f-d378-476f-a8bb-ed924c403fa8")
  const [sharing, setSharing] = useState("New and existing guest")
  const onSharingChange = e => {
    setSharing(e.target.value)
    e.target.value === "Anyone" ?
    setSharingId("0760ff30-1ac3-440e-9391-9c2e15533f33") :
    e.target.value === "New and existing guest" ?
    setSharingId("71d64c0f-d378-476f-a8bb-ed924c403fa8") :
    e.target.value === "Existing guest only" ?
    setSharingId("88cdcd9a-a286-4e1f-861c-6063dc873f0f") :
    setSharingId("3e060900-198c-4a1b-a6ad-5e4f597c4726")
  }


//SCRIPTS
/*
Id                  : 3e060900-198c-4a1b-a6ad-5e4f597c4726 OK
Title               :  External sharing Disabled (Team Site without Group)
WebTemplate         : 1
SiteScriptIds       : {721f126f-a657-4f38-8e44-4ddca33bb8be}
Description         : Sets External sharing to Disabled (Team Site without Group)

Id                  : 88cdcd9a-a286-4e1f-861c-6063dc873f0f
Title               :  External sharing ExistingExternalUserSharingOnly (Team Site without Group)
WebTemplate         : 1
SiteScriptIds       : {f5ce4b3c-7b29-44e5-9a8d-cdd8ad2db50b}
Description         : Sets External sharing to ExistingExternalUserSharingOnly (Team Site without Group)

Id                  : 71d64c0f-d378-476f-a8bb-ed924c403fa8
Title               :  External sharing ExternalUserSharingOnly (Team Site without Group)
WebTemplate         : 1
SiteScriptIds       : {6563274d-f5fe-451d-a916-f91e488c86eb}
Description         : Sets External sharing to ExternalUserSharingOnly (Team Site without Group)

Id                  : 0760ff30-1ac3-440e-9391-9c2e15533f33 OK
Title               :  External sharing ExternalUserAndGuestSharing (Team Site without Group)
WebTemplate         : 1
SiteScriptIds       : {3897ba25-22bd-40ad-9fb3-a2df5132c928}
Description         : Sets External sharing to ExternalUserAndGuestSharing (Team Site without Group)

*/


//SITE CREATION
  const createSite = async (e) => {
    e.preventDefault()
     const siteProps = {
      Owner: `${owners[0]}`,
      Title: `${domain}-${title}`,
      Url: `https://msfintl.sharepoint.com/sites/${domain}-${title}`,
      WebTemplate: "STS#3",
      SiteDesignId: `${sharingId}`
    };

    if (hub !== "") {
      siteProps['hubSiteId'] = `${hub}`
    }

    if (siteDesign !== "") {
      siteProps['siteDesignId']
    }
    // Log the current operation
    setProgress("Creating Team site ...");
  
    try {
        await sp.site.createCommunicationSiteFromProps(siteProps);
    } catch (error) {
        console.log(`Creating site error: ${error}`)
    }

    siteExists(`https://msfintl.sharepoint.com/sites/${domain}-${title}`)
  };

  const siteExistsChecker = async(titlecheck) => {
      try {
        const exists = await sp.site.exists(`https://msfintl.sharepoint.com/sites/${domain}-${titlecheck}`)
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
      
      setProgress("Team site created. Preparing other settings ...");
      await new Promise((resolve) => setTimeout(resolve, 10000));
      siteDesign === "" ? setProgress("Team site created") :
      designChecker ? await applyScript(siteUrl) : null
      owners.length !== 0 && await addSiteOwners(siteUrl)
      members.length !== 0 && await addSiteMembers(siteUrl)
      visitors.length !== 0 && await addSiteVisitors(siteUrl)
  }

  const applyScript = async(siteUrl) => {
        setProgress("Applying site design ...");
        const newsp = spfi(siteUrl).using(SPFxsp(context))
        try {
            await newsp.siteDesigns.applySiteDesign(
              `${siteDesign}`,
              `${siteUrl}`
            );
            setProgress("Other settings and scripts applied")
          } catch (error) {
            setError(`Error when applying site design: ${error}`);
          }
  } 
 
  const addSiteOwners = async(siteUrl) => {
    setProgress("Adding owners ...");
    
    const site = Web([sp.web, `${siteUrl}`]);
    owners.forEach((user) => {
      addSiteOwner(user,site)
    })
    setProgress("Owners added ...")
  }

    const addSiteOwner = async(user, site) =>{
          try {    
            const usersVisitors = await site.associatedOwnerGroup.users
            await usersVisitors.add(`${user}`);
          } catch (error) {
            setError(`Error when adding members: ${error}`);
          }
      }

  const addSiteMembers = async(siteUrl) => {
    setProgress("Adding members ...");
    
    const site = Web([sp.web, `${siteUrl}`]);
    members.forEach((user) => {
      addSiteMember(user,site)
    })
    setProgress("Members added ....")
  }

    const addSiteMember = async(user, site) =>{
          try {    
            const usersVisitors = await site.associatedMemberGroup.users
            await usersVisitors.add(`${user}`);
          } catch (error) {
            setError(`Error when adding members: ${error}`);
          }
      }

  const addSiteVisitors = async(siteUrl) => {
    setProgress("Adding visitors ...");  
    const site = Web([sp.web, `${siteUrl}`]);
    visitors.forEach((user) => {
      addSitevisitor(user,site)
    })
    setProgress("Visitors added ....")
  }

      const addSitevisitor = async(user, site) =>{
        try {    
          const usersVisitors = await site.associatedVisitorGroup.users
          await usersVisitors.add(`${user}`)
        } catch (error) {
          setError(`Error when adding users: ${error}`);
        }
      }


      const denyScripts = async(siteUrl) => {
        setProgress("Setting denyScript ...");

        const site = Web([sp.web, `${siteUrl}`]);
        try {    
          const status = await site()
          console.log(status)
    
        } catch (error) {
          setError(`Error when changing script: ${error}`);
        }
        
      }


  const disabled = title === "" || titleExist || !hubChecker || !designChecker ? true : false

  return (
    <div className={styles.site_wrapper}>
      <form className={styles.form_wrapper} onSubmit={createSite}>
          <label htmlFor='siteTitle'>Site name</label>
          <input id="siteTitle" type="text" onChange={addTitle}/>
          <span className={styles.input_comment}>{title === "" ? "Type a site title" : titleExist ? `NG, ${domain}-${title} already exists!` : "OK"}</span>
          <span className={styles.group_header}>Site users</span>
          <span>Site Owners</span>
          <PeoplePicker defaultSelectedUserIds={adminId} selectionMode="multiple" selectionChanged={addOwners}/>
          <span>Site Members</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addMembers}/>
          <span>Site Visitors</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addVisitors}/>
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
          <p>You will create a site without M365 group. Your site will have the following properties:</p>
          <h3>{domain}-{title}</h3>
          <span>Url: https://msfintl.sharepoint.com/sites/{domain}-{title}</span>
          <span>Sharing: {sharing}</span>
          <span>Site design: {siteDesignTitle}</span>
          <span>Associate with hub: {hubTitle}</span>
        </div>
        <div className={styles.result_progress}>
          {progress === "Finished" ? 
          <a target="_blank" href={`https://msfintl.sharepoint.com/sites/${domain}-${title}`}>Finished - click to open</a> :
          <span>{progress}</span>
          }
          {error}      
        </div>
      </div>
    </div>
  )
}

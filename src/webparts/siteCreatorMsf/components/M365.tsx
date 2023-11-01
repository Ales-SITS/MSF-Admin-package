import * as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';
import {  PeoplePicker } from '@microsoft/mgt-react';

//API
import { MSGraphClientV3  } from "@microsoft/sp-http";

//PNP SP
import { spfi, SPFx as SPFxsp} from "@pnp/sp";
import { Web } from "@pnp/sp/webs";  
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

//FLUENT
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

//Component
import Loader from './Loader'

export default function M365 (props) {

  //API
  const context = props.context
  const sp = spfi().using(SPFxsp(context))
  const graph = graphfi().using(SPFx(context))

  //General
  const domain = props.domain

  const [progress, setProgress] = useState("Not run yet")
  const [error, setError] = useState ("")
  const [title,setTitle] = useState(`GRP-${domain}-`)
  const [titleExist, setTitleExist] = useState(false)
  const [namingconventions,setNamingconventions] = useState(true)
  const [userdomain,setUserdomain] = useState(true)

  const addTitle = (e) => {
    setProgress("Not run yet")
    setError("")
    const title = e.target.value.replaceAll(" ","_")
    if (title.match(/^GRP-.*/)) {
      setTitle(title);
    }
    if (!title.match(/^GRP-.+-.*$/)){
      setNamingconventions(false)
    } else {setNamingconventions(true)}

    if (!title.startsWith(`GRP-${domain}`)){
      setUserdomain(false)
    } else {
      setUserdomain(true)
    }

    //setTitle(title)
    siteExistsChecker(title)
  }

//Content types
const ct = props.ct
const ct_options: IDropdownOption[] = [...ct]
const [selected_ct,setSelected_ct] = useState([])
const selectedHandler_ct = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
  console.log(item)
  if (item.selected) {
    setSelected_ct([...selected_ct, {key: item.key, text: item.text}]);
  } else {
    setSelected_ct(selected_ct.filter(item => item.key !== item.key));
  }
};

//Avaiable Hubsites
const hs = props.hs
const hs_options: IDropdownOption[] = [{key:null, text: "none"}, ...hs]
const [selected_hs,setSelected_hs] = useState({key: null, text: "none"})
const selectedHandler_hs = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    setSelected_hs(item ? item : undefined);
};


//Avaiable Site designs
const sd = props.sd
const sd_options: IDropdownOption[] = [{key:null, text: "none"}, ...sd]
const [selected_sd,setSelected_sd] = useState({key: null, text: "none"})
const selectedHandler_sd = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
  setSelected_sd(item ? item : undefined);
};

  //Users states
  const[adminId,setAdminId] =useState([])
  const [owners,setOwners] = useState([])
  const [members,setMembers] = useState([])
  const [siteOwners,setSiteOwners] = useState([])
  const [siteMembers,setSiteMembers] = useState([])
  const [siteVisitors,setSiteVisitors] = useState([])

  const addOwners = (e) => {
    const ownersMails = e.detail.map(owner => `https://graph.microsoft.com/v1.0/users/${owner.id}`)
    setOwners([...ownersMails])
  }

  const addMembers = (e) => {
    const membersMails = e.detail.map(member => `https://graph.microsoft.com/v1.0/users/${member.id}`)
    setMembers(membersMails)
  }

  const addSiteOwners = (e) => {
    const owners = e.detail.map(owner =>  `i:0#.f|membership|${owner.userPrincipalName.toLowerCase()}`)
    setSiteOwners([...owners])
  }

  const addSiteMembers = (e) => {
    const members = e.detail.map(member => `i:0#.f|membership|${member.userPrincipalName.toLowerCase()}`)
    setSiteMembers(members)
  }

  const addSiteVisitors = (e) => {
    const visitors = e.detail.map(member => `i:0#.f|membership|${member.userPrincipalName.toLowerCase()}`)
    setSiteVisitors(visitors)
  }

  //Admin check
  useEffect(()=>{
    getAdmin().then(result => {
      setAdminId([result.id])
  })
  },
  [])

  //Privacy state
  const [privacy, setPrivacy] = useState("Private")

  const onPrivacyChange = e => {
    setPrivacy(e.target.value)
  }

  //Sharing state
  const [sharingId, setSharingId] = useState("6bf7f2ba-2f3f-4449-9c7d-596f22e9cb83")
  const [sharing, setSharing] = useState("New and existing guest")
  const onSharingChange = e => {
    setSharing(e.target.value)
    e.target.value === "Anyone" ?
    setSharingId("004bbf35-ed4a-45e6-9199-cc9881aeba64") :
    e.target.value === "New and existing guest" ?
    setSharingId("6bf7f2ba-2f3f-4449-9c7d-596f22e9cb83") :
    e.target.value === "Existing guest only" ?
    setSharingId("513beb79-9027-4297-a38d-db7cbfb83b07") :
    setSharingId("3b0ffe11-9840-4561-96a9-c6c417976db9")
  }

//SITE CREATION
  const createSite = (e) => {
    e.preventDefault()
     const group = {
      description: 'For API testing purposes',
      displayName: `${title}`,
      groupTypes: ['Unified'],
      mailEnabled: true,
      securityEnabled: false,
      visibility: `${privacy}`,
      mailNickname: `${title}`
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
          siteExists(`https://msfintl.sharepoint.com/sites/${title}`)      
        })
      .catch((error: any) => {
          setError(`Error when creating site: ${error}`);
      });
      })    
  };

  const siteExistsChecker = async(titlecheck) => {
      try {
        const exists = await sp.site.exists(`https://msfintl.sharepoint.com/sites/${titlecheck}`)
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
      if (selected_ct.length > 0) { 
        await applyTaxonomy(siteUrl)
        for (const ct of selected_ct) {
          setProgress(`Adding ${ct.text} content type ...`)
          await includeContentTypes(ct.key,siteUrl)
        }  
        } 

      await applyScript(siteUrl,sharingId, 1);
      siteOwners.length !== 0 && await addSiteOwnersCall(siteUrl)
      siteMembers.length !== 0 && await addSiteMembersCall(siteUrl)
      siteVisitors.length !== 0 && await addSiteVisitorsCall(siteUrl)
      selected_sd.key !== null ? await applyScript(siteUrl,selected_sd.key, 0) : null
      selected_hs.key !== null ? await associateToHub(siteUrl) : null
      setProgress("Finished")
  }

  const applyTaxonomy = async(siteUrl) => {
    const newsp = spfi(siteUrl).using(SPFxsp(context))
    try {
      await newsp.site.features.add("73EF14B1-13A9-416b-A9B5-ECECA2B0604C")
      setProgress("Adding taxonomy feature")
    } catch (error) {
      console.log(`Error when adding taxonomy feauture: ${error}`)
    }

  }

  const applyScript = async(siteUrl,designId,type) => {
    type === 1 ? setProgress("Applying external sharing settings ...") : setProgress("Applying site design ...")
    const newsp = spfi(siteUrl).using(SPFxsp(context))
    try {
        await newsp.siteDesigns.applySiteDesign(
          `${designId}`,
          `${siteUrl}`
        );
        type === 1 ? setProgress("External sharing set ...") : setProgress("Site design applied ...")
      } catch (error) {
        type === 1 ? setError(`Error when setting External sharing: ${error}`) : setError(`Error when applying site design: ${error}`)
      }
} 

  async function associateToHub (siteUrl) {
      setProgress("Associating with hub ...");
        const newsp = spfi(siteUrl).using(SPFxsp(context))
        try {
          await newsp.site.joinHubSite(`${selected_hs.key}`)
          setProgress("Finished")
        } catch (error) {
          setError(`Error when associating to hub: ${error}`);
        }
  }

  async function getAdmin(){
    
    const admin = await graph.me()

    return admin
  }

  const addSiteOwnersCall = async(siteUrl) => {
    setProgress("Adding site owners ...");
    
    const site = Web([sp.web, `${siteUrl}`]);
    siteOwners.forEach((user) => {
      addSiteOwner(user,site)
    })
    setProgress("Site owners added ...")
  }

    const addSiteOwner = async(user, site) =>{
          try {    
            const usersVisitors = await site.associatedOwnerGroup.users
            await usersVisitors.add(`${user}`);
          } catch (error) {
            setError(`Error when adding members: ${error}`);
          }
      }

  const addSiteMembersCall = async(siteUrl) => {
    setProgress("Adding site members ...");
    
    const site = Web([sp.web, `${siteUrl}`]);
    siteMembers.forEach((user) => {
      addSiteMember(user,site)
    })
    setProgress("Site members added ...")
  }

    const addSiteMember = async(user, site) =>{
          try {    
            const usersVisitors = await site.associatedMemberGroup.users
            await usersVisitors.add(`${user}`);
          } catch (error) {
            setError(`Error when adding members: ${error}`);
          }
      }

  const addSiteVisitorsCall = async(siteUrl) => {
    setProgress("Adding site visitors ...");  
    const site = Web([sp.web, `${siteUrl}`]);
    siteVisitors.forEach((user) => {
      addSitevisitor(user,site)
    })
    setProgress("Site visitors added ....")
  }

      const addSitevisitor = async(user, site) =>{
        try {    
          const usersVisitors = await site.associatedVisitorGroup.users
          await usersVisitors.add(`${user}`)
        } catch (error) {
          setError(`Error when adding users: ${error}`);
        }
      }

    async function includeContentTypes (id,siteURL) {
        const urlObject = new URL(siteURL);
        const host = urlObject.hostname
        const path = urlObject.pathname
        try {
          await graph.sites.getByUrl(host, path).contentTypes.addCopyFromContentTypeHub(`${id}`);
        } catch (error) {
          setError(`Error when syncing content types: ${error}`);
        } 
    }

  const disabled = title === "" || titleExist || namingconventions === false ? true : false

  return (
    <div className={styles.site_wrapper}>
      <form className={styles.form_wrapper} onSubmit={createSite}>
          <label htmlFor='siteTitle' className={styles.group_header}>Site name <span className={styles.input_hint}>{`(GRP-DOMAIN-Purpose)`}</span></label>
          <input 
            id="siteTitle" 
            type="text" 
            value={`${title}`} 
            pattern="GRP-.*" 
            onChange={addTitle}
          />
          <span className={styles.input_comment}>
            {
             title === `GRP-${domain}-` ? `Type a site title in the "GRP-DOMAIN-Purpose" format` : 
             titleExist ? `NG, ${title} already exists` : 
             !namingconventions ? `NG, please follow the "GRP-DOMAIN-Purpose" naming convention!` :            
             "OK"}</span>
          <span className={styles.input_comment}>{userdomain? null : "⚠ Site domain part doesn't match your account domain!"}</span>

          <span className={styles.group_header}>M365 Group users</span>
          <span>M365 group Owners</span>
          <PeoplePicker defaultSelectedUserIds={adminId} selectionMode="multiple" selectionChanged={addOwners}/>
          <span>M365 group Members</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addMembers}/>
          <span className={styles.group_header}>Site only users</span>
          <span className={styles.group_header_comment}>M365 Group users will be added automatically</span>
          <span>Site Owners</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addSiteOwners}/>
          <span>Site Members</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addSiteMembers}/>
          <span>Site Visitors</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addSiteVisitors}/>
          <div className={styles.selection_box_wrapper}>
            <div className={styles.selection_box}>
            <span className={styles.group_header}>Privacy <a className={styles.help_link} target="_blank" rel="noreferrer" href="https://www.mrsharepoint.guru/microsoft-office-365-groups-private-vs-public">?</a></span>
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
              <span className={styles.group_header}>External sharing <a className={styles.help_link} target="_blank" rel="noreferrer" href="https://learn.microsoft.com/en-US/sharepoint/change-external-sharing-site?WT.mc_id=365AdminCSH_inproduct#which-option-to-select">?</a></span>
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
          </div>
          <span className={styles.group_header}>Other</span>
          <Dropdown
            placeholder="Select"
            label="Associate to hub"
            defaultSelectedKey={selected_hs.key}
            options={hs_options}
            onChange={selectedHandler_hs}
          />
          <Dropdown
            placeholder="Select"
            label="Apply site design"
            defaultSelectedKey={selected_sd.key}
            options={sd_options}
            onChange={selectedHandler_sd}
          />
          <Dropdown
            placeholder="Select"
            label="Select content type(s), if handled in site designs and flows"
            multiSelect
            options={ct_options}
            onChange={selectedHandler_ct}
          />
          <div className={styles.createSite_button_wrapper}>
            <input className={styles.createSite_button} 
                   type="submit" onClick={createSite} 
                   value="Create site" 
                   disabled = {disabled}/>
          </div>
      </form>
      <div className={styles.result_wrapper}>
      <div className={styles.result_list}>
          <p>You will create a team site with M365 group. Your site will have the following properties:</p>
          <h3>{title}</h3>
          <div className={styles.result_list_details}>
            <span>Url:</span>
            <span>https://msfintl.sharepoint.com/sites/{title}</span>

            <span>Privacy:</span>
            <span>{privacy}</span>

            <span>Sharing:</span>
            <span>{sharing}</span>

            <span>Associated with hub:</span>
            <span>{selected_hs.text}</span>

            <span>Site design:</span>
            <span>{selected_sd.text}</span> 

            <span>Content types:</span>
            <div className={styles.result_list_ct}>
              {selected_ct.map(ct => <span>{ct.text}</span>)}
            </div>
          </div>
      </div>
      <div className={styles.result_progress}>
          <span className={styles.error_message}>{error}</span> 
          {progress === "Finished" ? 
          <a className={styles.finished_link} target="_blank" href={`https://msfintl.sharepoint.com/sites/${title}`}><span className={styles.finished_link_check}>&#10003;</span> Finished - click to open</a> :
          <span>{progress}</span>
          }
          {progress !== "Finished" && progress !== "Not run yet" ? <Loader/> : null}
    </div>
      </div>
    </div>
  )
}

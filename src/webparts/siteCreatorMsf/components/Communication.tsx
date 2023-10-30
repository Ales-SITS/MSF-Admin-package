import * as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';
import {  PeoplePicker } from '@microsoft/mgt-react';

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
import "@pnp/graph/sites";
import "@pnp/graph/content-types";

//FLUENT
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

//Component
import Loader from './Loader'

export default function Communication (props) {

  //API
  const context = props.context
  const sp = spfi().using(SPFxsp(context))
  const graph = graphfi().using(SPFx(context))

  //General
  const domain = props.domain

  const [progress, setProgress] = useState("Not run yet")
  const [error, setError] = useState ("")
  const [title,setTitle] = useState(`${domain}-`)
  const [titleExist, setTitleExist] = useState(false)
  const [namingconventions,setNamingconventions] = useState(true)
  const [userdomain,setUserdomain] = useState(true)

  //Content types
  const ct = props.ct
  const ct_options: IDropdownOption[] = [...ct]
  const [selected_ct,setSelected_ct] = useState([])
  const selectedHandler_ct = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
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


 //Sharing state
  const [sharingId, setSharingId] = useState("1de7f8a5-b635-42a9-b8c0-32ae09e26765")
  const [sharing, setSharing] = useState("New and existing guest")
  const onSharingChange = e => {
    setSharing(e.target.value)
    e.target.value === "Anyone" ?
    setSharingId("6ddcd576-4e03-449c-bdf6-fa555a460674") :
    e.target.value === "New and existing guest" ?
    setSharingId("1de7f8a5-b635-42a9-b8c0-32ae09e26765") :
    e.target.value === "Existing guest only" ?
    setSharingId("9251bc06-7392-4889-bf50-275a42d63699") :
    setSharingId("8759e9ce-6309-4d33-b499-0c06dbc141a0")
  }

//SITE CREATION
  const createSite = async (e) => {
    e.preventDefault()
     const siteProps = {
      Owner: `${owners[0]}`,
      Title: `${title}`,
      Url: `https://msfintl.sharepoint.com/sites/${title}`,
      WebTemplate: "SITEPAGEPUBLISHING#0"    
    };

    setProgress("Creating Communication site ...");
  
    try {
        await sp.site.createCommunicationSiteFromProps(siteProps);
    } catch (error) {
        console.log(`Creating site error: ${error}`)
    }

    siteExists(`https://msfintl.sharepoint.com/sites/${title}`)
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
      
      setProgress("Communication site created. Preparing other settings ...");
      await new Promise((resolve) => setTimeout(resolve, 10000));
      
    
      if (selected_ct.length > 0) { 
      await applyTaxonomy(siteUrl)
      for (const ct of selected_ct) {
        setProgress(`Adding ${ct.text} content type ...`)
        await includeContentTypes(ct.key,siteUrl)
      }  
      } 
      await applyScript(siteUrl,sharingId, 1)
      owners.length !== 0 && await addSiteOwners(siteUrl)
      members.length !== 0 && await addSiteMembers(siteUrl)
      visitors.length !== 0 && await addSiteVisitors(siteUrl)
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
    setProgress("Members added ...")
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
    setProgress("Visitors added ...")
  }

      const addSitevisitor = async(user, site) =>{
        try {    
          const usersVisitors = await site.associatedVisitorGroup.users
          await usersVisitors.add(`${user}`)
        } catch (error) {
          setError(`Error when adding users: ${error}`);
        }
      }

  async function associateToHub (siteUrl) {
        setProgress("Associating with hub ...");
          const newsp = spfi(siteUrl).using(SPFxsp(context))
          try {
            await newsp.site.joinHubSite(`${selected_hs.key}`)
            setProgress("Associated with the hub ...")
          } catch (error) {
            setError(`Error when associating to hub: ${error}`);
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

  const disabled = title === "" || titleExist ? true : false

  const addTitle = (e) => {
    setProgress("Not run yet")
    setError("")
    const title = e.target.value.replaceAll(" ","_")
    setTitle(title)
    siteExistsChecker(e.target.value)

    if (!title.match(/.+-.+/)){
      setNamingconventions(false)
    } else {setNamingconventions(true)}

    if (!title.startsWith(`${domain}`)){
      setUserdomain(false)
    } else {
      setUserdomain(true)
    }

  }


  return (
    <div className={styles.site_wrapper}>
      <form className={styles.form_wrapper} onSubmit={createSite}>
          <label htmlFor='siteTitle' className={styles.group_header}>Site name <span className={styles.input_hint}>{`(DOMAIN-Purpose)`}</span></label>
          <input 
          id="siteTitle" 
          type="text" 
          value={`${title}`}  
          onChange={addTitle}
          />
          <span className={styles.input_comment}>
            {title === `${domain}-` ? `Type a site title in the "DOMAIN-Purpose" format` : titleExist ? `NG, ${title} already exists!` : "OK"}
          </span>
          {namingconventions? null : <span className={styles.input_comment}> "⚠ Please follow the "DOMAIN-Purpose" naming convention!"</span>}
          {userdomain? null : <span className={styles.input_comment}> "⚠ Site domain part doesn't match your account domain!"</span>}
          <span className={styles.group_header}>Site users</span>
          <span>Site Owners</span>
          <PeoplePicker defaultSelectedUserIds={adminId} selectionMode="multiple" selectionChanged={addOwners}/>
          <span>Site Members</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addMembers}/>
          <span>Site Visitors</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addVisitors}/>
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
            <input className={styles.createSite_button} type="submit" onClick={createSite} value="Create site" 
                   disabled = {disabled}/>
          </div>
      </form>
      <div className={styles.result_wrapper}>
        <div className={styles.result_list}>
          <p>You will create a communication site. Your site will have the following properties:</p>
          <h3>{domain}-{title}</h3>
          <div className={styles.result_list_details}>
            <span>Url:</span>
            <span>https://msfintl.sharepoint.com/sites/{title}</span>

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

import * as React from 'react';
import {useState, useEffect} from 'react'
import styles from './SiteCreatorMsf.module.scss';
import { ISiteCreatorMsfProps } from './ISiteCreatorMsfProps';
import {  PeoplePicker } from '@microsoft/mgt-react';

import { MSGraphClientV3  } from "@microsoft/sp-http";

import { spfi, SPFx as SPFxsp} from "@pnp/sp";
import "@pnp/sp/site-designs"
import "@pnp/sp/sites";


//GRAPH
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";


export default function SiteCreatorMsf (props) {

  const context = props.context
  const sp = spfi().using(SPFxsp(context))
  const graph = graphfi().using(SPFx(context))

  const domain = context.pageContext.user.email?.split("@")[1]?.split(".")[0]?.toUpperCase()

  const [title,setTitle] = useState("")
  const[admin,setAdmin] =useState([])
  const [owners,setOwners] = useState([])
  const [members,setMembers] = useState([])
  const [siteCreated,setSiteCreated] = useState(false)

  const addTitle = (e) => {
    setTitle(e.target.value)
  }

  const addOwners = (e) => {
    const ownersMails = e.detail.map(owner => `https://graph.microsoft.com/v1.0/users/${owner.id}`)
    setOwners([...admin,...ownersMails])
  }

  const addMembers = (e) => {
    const membersMails = e.detail.map(member => `https://graph.microsoft.com/v1.0/users/${member.id}`)
    setMembers(membersMails)
  }

  const [privacy, setPrivacy] = useState("private")

  const onOptionChange = e => {
    setPrivacy(e.target.value)
  }

// Site script externalsharing disabled 
// de32b387-f7af-4d9f-8e82-40ad7fa23500
// Site design (team site)
// 123ac0ed-b076-4507-82e4-de444923a4b5


  const sharingSetting = async() => {
      const test = await sp.siteDesigns.applySiteDesign("123ac0ed-b076-4507-82e4-de444923a4b5",`https://msfintl.sharepoint.com/sites/GRP-${domain}-${title}`)
      console.log(test)
    // Log the current operation
    /*
    console.log("Reading Site...739ffdd2-6522-4566-b0e9-38d67c3fe426");
      
    props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        client.api('/sites/739ffdd2-6522-4566-b0e9-38d67c3fe426')
	      .version('beta')
	      .get();   
          });*/
  };


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
    console.log("Creating Site...");
  
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


  const siteExists = async(siteUrl) => {
      let exists = false

      while(!exists) {
        exists =  await sp.site.exists(siteUrl)
        if(!exists){
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }
      setSiteCreated(exists)

      console.log("BOO")

      await new Promise((resolve) => setTimeout(resolve, 10000));

      console.log("RUN")

      await applyScript(siteUrl);
  }

  const applyScript = async(siteUrl) => {
    console.log(siteUrl)
        try {
            const script = await sp.siteDesigns.applySiteDesign(
              "123ac0ed-b076-4507-82e4-de444923a4b5",
              `${siteUrl}`
            );
          } catch (error) {
            console.error('Error applying site design:', error);
          }
   }

    

    async function getAdmin():Promise<any>{
    
    const admin = await graph.me()

    return admin
}


   useEffect(()=>{
    getAdmin().then(result => {
      setAdmin([`https://graph.microsoft.com/v1.0/users/${result.id}`])
   })
  },
   [])

  console.log(owners)

  return (
    <div className={styles.m365_wrapper}>
      <form className={styles.form_wrapper} onSubmit={createSite}>
          <span>Site name</span>
          <input type="text" onChange={addTitle}/>
          <span>Owners</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addOwners}/>

          <span>Members</span>
          <PeoplePicker selectionMode="multiple" selectionChanged={addMembers}/>
          <input
              type="radio"
              name="privacy"
              value= "private"
              id="private"
              checked={privacy === "private"}
              onChange={onOptionChange}
          />
          <label htmlFor="private">Private</label>
          <input
              type="radio"
              name="privacy"
              value= "public"
              id="public"
              checked={privacy === "public"}
              onChange={onOptionChange}
          />
          <label htmlFor="public">Public</label>
          <input type="submit" onClick={createSite} value="Create site"/>
      </form>
      <div className={styles.result_wrapper}>
          <h3>GRP-{domain}-{title}</h3>
          <span>Url: https://msfintl.sharepoint.com/sites/GRP-{domain}-{title}</span>
          <span>{privacy}</span>
          <span>Site created: {`${siteCreated}`}</span>
          {siteCreated && <button onClick={sharingSetting}>Change sharing settigs</button>}
      </div>
    </div>
  )
}

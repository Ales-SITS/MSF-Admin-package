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

export default function SiteCreatorMsf (props) {

  const graph = graphfi().using(SPFx(props.context))

  const [selectedType,setSelectedType] = useState(1)
  const selectedTypeHandler = (type) => {
    setSelectedType(type)
  }

  const [isAdmin, setIsAdmin] = useState(false)
  const [loader,setLoader] = useState(true)

  useEffect(()=>{
    getAdminGroup()
  },[])

  const getAdminGroup = async() => {
    const admin = await graph.me();
    const admingroups = await graph.me.memberOf();
    const allowgroup = await graph.groups.getById("d35fc400-84f4-40c1-956f-c0532a976d1f").members();
    for (const grpA of admingroups) {
      if (allowgroup.some(grpB => grpA.id === grpB.id)) {
        setIsAdmin(true)
      }   
    }
    setLoader(false)
  }

  return (
    <div className={styles.app_wrapper}>
      <h2>SITE Creator</h2>
      <div className={styles.app_navigation}>
        <button onClick={()=>selectedTypeHandler(1)} className={selectedType === 1 && styles.selected_site_type}>Teams site (M365)</button>
        <button onClick={()=>selectedTypeHandler(2)} className={selectedType === 2 && styles.selected_site_type}>Team site without M365</button>
        <button onClick={()=>selectedTypeHandler(3)} className={selectedType === 3 && styles.selected_site_type}>Communication site</button>
      </div>
      {!isAdmin ? <Declined context={props.context} loader={loader}/> :
       selectedType === 1 ? <M365 context={props.context}/> :
       selectedType === 2 ? <NonM365 context={props.context}/> : 
       <Communication context={props.context}/>
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
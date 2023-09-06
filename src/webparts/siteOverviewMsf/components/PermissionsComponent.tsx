import * as React from 'react';
import {useState, useEffect, useRef} from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPFx, graphfi } from "@pnp/graph";

import styles from './SiteOverviewMsf.module.scss';
import perstyles from './PermissionsComponent.module.scss';

import { spfi, SPFx as SPFxsp} from "@pnp/sp";

import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/security";

import "@pnp/graph/users";
import "@pnp/graph/sites";

import { Web } from "@pnp/sp/webs";   

import { Icon } from '@fluentui/react/lib/Icon';

import { PermissionKind } from "@pnp/sp/security";
import { PropertyPaneDescription } from 'SiteOverviewMsfWebPartStrings';

export default function PermissionsComponent (props) {

    const [everyone, setEveryone] = useState(false)
    const [users,setUsers] = useState([])
    async function getUsers() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.siteUsers()    

        return users
    }
    
    const [userGroup,setUserGroup] = useState([])
    async function getUserGroup(id) {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.siteUsers.getById(id).groups()    
    
        return users
    } 

    const [groups,setGroups] = useState([])
    async function getGroups() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const userGroups = await site.siteGroups()
        return userGroups
    }  


    const [role,setRole] = useState(null)
    async function getRole() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const web = Web([sp.web, `${props.url}`])      
        const users = await web.roleAssignments()
    
        return users
    }  


    useEffect(() => {
        getUsers().then(result => {
            setUsers([]);
            const arr:any = result.filter(user => user.Email!=="")
            setUsers(arr);

            const every:any = result.filter(user => user.LoginName.startsWith("c:0-.f"))
     
            every.length > 0 ? setEveryone(true) : setEveryone(false)

            getUserGroup(arr[0].Id).then(result => {
                setUserGroup(result)
    
            })
        });
 
        getRole().then(result => {
            setRole(result)
        })

        }, [props]);

        //console.log(role)
//VISUAL
const closeHandler = () => {
    props.onCloseHandler()
}

//console.log(role)
//FILTERED
const [displayCount, setDisplayCount] = useState(30);

const [usersFilter,setUsersFilter] = useState("")
const usersFilterHandler = (e) => {
    setUsersFilter(e)
    setDisplayCount(50)
}

const usersfiltered = usersFilter === "" ? users : users.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase()))
const usersToDisplay = usersfiltered.slice(0, displayCount);

const handleScroll = (event) => {
    event.currentTarget.scrollTop + event.currentTarget.offsetHeight >= event.currentTarget.scrollHeight ? setDisplayCount(displayCount + 30) : null
}

//console.log(usersToDisplay)

return (
        <div className={perstyles.permissionModal}>
        <div className={perstyles.permissionModalTop}>
            <div>
                <h2>Permissions overview</h2>
                <span>for {props.url}</span>
            </div>
            <button onClick={closeHandler}><Icon iconName="ChromeClose"/></button>
        </div>
        <div className={perstyles.permissionModalInput}>
            <input 
                type="text" 
                name="siteName" 
                placeholder="Filter by user name"
                onChange={e => usersFilterHandler(e.target.value)} 
                /><span>({displayCount}/{usersfiltered.length})</span>
        </div>
        <div className={perstyles.permissionModalResults} onScroll={handleScroll}>
            <ul>
            {usersToDisplay.map((user,idx) =>
            <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
            )}
            </ul>
        </div>
    </div>
        
    );
  }


  function PersonPermissions (props) {
    const user = props.user
    //console.log(props.user.LoginName)
    const [up,setUp] = useState(null)
    async function getUp(loginName) {  
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const user = await site.getUserEffectivePermissions(loginName) 
        
        console.log(user.Low)

        //High: '2147483647', Low: '4294967295'  Site Admin
        //High: '2147483647', Low: '4294967295'

        //High: '432', Low: '1011030767'  Edit
        //High: '176', Low: '138612833'  Read

        //High: '0', Low: '0' access without access

        /*High Value: The "High" value represents the permissions that 
        are explicitly granted to the user or group on the resource.
        These are the permissions that have been assigned directly to 
        the user or group through SharePoint permissions management.

        Low Value: The "Low" value represents the permissions that are granted through group memberships.
        It includes the permissions inherited from SharePoint groups that the user is a member of. 
        For example, if a user is a member of a SharePoint group that has read access to a document library, 
        the "Low" value will include these permissions.
        */

        if (user.Low === 4294967295) {return "Full control"}
        else if (user.Low === 1011030767) {return "Edit"}
        else if (user.Low === 138612833) {return "Read"}
        else if (user.Low === 0) {return "No access"}
        /*
        if(
            //FULL
            sp.web.hasPermissions(user,PermissionKind.FullMask)        
        ) {
            console.log(`${user} - is admin`)
            return "Full control permissions"
        } else if  (
            //EDIT
            sp.web.hasPermissions(user,PermissionKind.ManageLists ) &&
            sp.web.hasPermissions(user,PermissionKind.AddListItems) &&
            sp.web.hasPermissions(user,PermissionKind.EditListItems) &&
            sp.web.hasPermissions(user,PermissionKind.EditListItems) &&
            sp.web.hasPermissions(user,PermissionKind.ViewListItems) &&
            sp.web.hasPermissions(user,PermissionKind.OpenItems) &&
            sp.web.hasPermissions(user,PermissionKind.ViewVersions) &&
            sp.web.hasPermissions(user,PermissionKind.DeleteVersions) &&
            sp.web.hasPermissions(user,PermissionKind.CreateAlerts) &&
            sp.web.hasPermissions(user,PermissionKind.ViewFormPages) &&
            sp.web.hasPermissions(user,PermissionKind.BrowseDirectories) &&
            sp.web.hasPermissions(user,PermissionKind.CreateSSCSite) &&
            sp.web.hasPermissions(user,PermissionKind.ViewPages) &&
            sp.web.hasPermissions(user,PermissionKind.BrowseUserInfo) &&
            sp.web.hasPermissions(user,PermissionKind.UseRemoteAPIs) &&
            sp.web.hasPermissions(user,PermissionKind.UseClientIntegration) &&
            sp.web.hasPermissions(user,PermissionKind.Open) &&
            sp.web.hasPermissions(user,PermissionKind.EditMyUserInfo) &&
            sp.web.hasPermissions(user,PermissionKind.ManagePersonalViews) &&
            sp.web.hasPermissions(user,PermissionKind.AddDelPrivateWebParts) &&
            sp.web.hasPermissions(user,PermissionKind.UpdatePersonalWebParts)
        ) {
            return "Edit permissions"
        } else if (
            //READ
            sp.web.hasPermissions(user,PermissionKind.ViewListItems) &&
            sp.web.hasPermissions(user,PermissionKind.OpenItems) &&
            sp.web.hasPermissions(user,PermissionKind.ViewVersions) &&
            sp.web.hasPermissions(user,PermissionKind.CreateAlerts) &&
            sp.web.hasPermissions(user,PermissionKind.ViewFormPages) &&
            sp.web.hasPermissions(user,PermissionKind.CreateSSCSite) &&
            sp.web.hasPermissions(user,PermissionKind.ViewPages) &&
            sp.web.hasPermissions(user,PermissionKind.BrowseUserInfo) &&
            sp.web.hasPermissions(user,PermissionKind.UseRemoteAPIs) &&
            sp.web.hasPermissions(user,PermissionKind.UseClientIntegration) &&
            sp.web.hasPermissions(user,PermissionKind.Open)
        ) {
            return "Read permissions"
        } else {
            return "Other permissions"
        }
        */
        return user.Low.toString()
    }  

    useEffect(() => {
        getUp(props.user.LoginName).then(result => {
            
            result === '4294967295' ? setUp("Full control") :
            result === '1011030767' ? setUp("Edit") :
            result === '138612833' ? setUp("Read") :
            result === '0' ? setUp("No access") :
            setUp("Other")
        });
 
        },[props]);

    //console.log(user)
    
    return (
        <li>    
            <div className={perstyles.userPermBox}>
                <span>
                    {user.Title}{user.IsSiteAdmin && " (Site Admin)"}
                </span>
                <span>
                    {user.Email}
                </span> 
                <span>
                    {up}
                </span>  
            </div>
         
        </li>
    )
  }


  
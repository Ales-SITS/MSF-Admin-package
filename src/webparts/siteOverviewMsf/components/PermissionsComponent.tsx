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

console.log(role)
//FILTERED
const [displayCount, setDisplayCount] = useState(50);

const [usersFilter,setUsersFilter] = useState("")
const usersFilterHandler = (e) => {
    setUsersFilter(e)
    setDisplayCount(50)
}

const usersfiltered = users.filter( user => user.Title.includes(usersFilter))
const usersToDisplay = usersfiltered.slice(0, displayCount);

const handleScroll = (event) => {
    event.currentTarget.scrollTop + event.currentTarget.offsetHeight >= event.currentTarget.scrollHeight ? setDisplayCount(displayCount + 50) : null
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
                placeholder="Filter by site Title"
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
        const users = await site.getUserEffectivePermissions(loginName) 
        
        console.log(users)

        return users
    }  

    useEffect(() => {
        getUp(props.user.LoginName).then(result => {  

            setUp(result);
        });
 
        },[]);

   
    
    return (
        <li>    
            <div className={perstyles.userPermBox}>
                <span>
                    {user.Title}
                </span>
                <span>
                    {user.Email}
                </span> 
                <span>
                    {up === null ? null : `${(up.High >>> 0).toString(2)}/${(up.Low >>> 0).toString(2)}`}
                </span>  
            </div>
         
        </li>
    )
  }


  
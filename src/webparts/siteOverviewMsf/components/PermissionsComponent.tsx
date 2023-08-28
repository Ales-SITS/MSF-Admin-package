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

export default function PermissionsComponent (props) {

    const [users,setUsers] = useState([])
    async function getUsers() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.siteUsers()    
    
        return users
    }   

    const [groups,setGroups] = useState([])
    async function getGroups() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.siteGroups()   
    
        return users
    }  


    const [up,setUp] = useState(null)
    async function getUp() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.getCurrentUserEffectivePermissions() 
    
        return users
    }  


    useEffect(() => {
        getUsers().then(result => {
            setUsers([]);
            const arr:any = result.filter(per => !per.Title.startsWith("SharingLinks."))
            setUsers(arr);
        });
 
        getGroups().then(result => {
            setGroups([]);
            //const arr:any = result.filter(per => !per.Title.startsWith("SharingLinks."))
            setGroups(result);
        });

        getUp().then(result => {
            setUp(null);
            //const arr:any = result.filter(per => !per.Title.startsWith("SharingLinks."))
            setUp(result);
        });

        }, [props]);

    console.log(users)
    console.log(groups)
    console.log(up)

     return (
        <div className={perstyles.permissionBoxAnchor}>
            <div className={perstyles.permissionBox}>
                <div>Number of SPO users: {users.length}</div>
                <div>Number of SPO groups: {groups.length}</div>
                <div className={perstyles.permissions}>
                    <ul>
                    {/*users.map((user,idx) => 
                        <li key={idx}>{user.Email}</li>)*/}
                    </ul>
                    <ul>
                    {/*groups.map((group,idx) => 
                        <li key={idx}>{group.LoginName}</li>)*/}
                    </ul>
                </div>


            </div>
    
        </div>    
    );
  }

import * as React from 'react';
import {useState, useEffect} from 'react';

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


async function getGroups(web,url) {
    const site = Web([web, `${url}`]) 
    const groups = await site.siteGroups()
    return groups
 }

async function getOwners(web,url) {
    const site = Web([web, `${url}`]) 
    const groupid = await site.associatedOwnerGroup().then(result => {return result.Id})
    const owners = await site.siteGroups.getById(groupid).users();
    return owners
 }

 async function getMembers(web,url) {
    const site = Web([web, `${url}`]) 
    const groupid = await site.associatedMemberGroup().then(result => {return result.Id})
    const owners = await site.siteGroups.getById(groupid).users();
    return owners
 }

 async function getVisitors(web,url) {
    const site = Web([web, `${url}`]) 
    const groupid = await site.associatedVisitorGroup().then(result => {return result.Id})
    const owners = await site.siteGroups.getById(groupid).users();
    return owners
 }


export default function PermissionsComponent (props) {
    const sp = props.sp
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

    const [owners,setOwners] = useState([])
    const [members,setMembers] = useState([])
    const [visitors,setVisitors] = useState([])

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
       
        getOwners(sp.web,props.url).then(result => setOwners(result))
        getMembers(sp.web,props.url).then(result => setMembers(result))
        getVisitors(sp.web,props.url).then(result => setVisitors(result))

        //getGroups(sp.web,props.url).then(result => console.log(result.filter(group => !group.LoginName.includes("SharingLinks"))))
        //getRole()

        }, []);


//VISUAL
const closeHandler = () => {
    props.onCloseHandler()
}


//FILTERED
const [displayCount, setDisplayCount] = useState(30);

const [usersFilter,setUsersFilter] = useState("")
const usersFilterHandler = (e) => {
    setUsersFilter(e)
    setDisplayCount(50)
}


const admins = users.filter(user => user.IsSiteAdmin === true)
const adminfiltered = usersFilter === "" ? admins :
      admins.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
)

const ownersfiltered = usersFilter === "" ? owners :
owners.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
)
const membersfiltered = usersFilter === "" ? members :
members.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
)
const visitorsfiltered = usersFilter === "" ? visitors :
visitors.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
)


const internal = users.filter (user => user.IsShareByEmailGuestUser === false)
const internalfiltered = usersFilter === "" ? internal :
      internal.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
      )
const internalToDisplay = internalfiltered.slice(0, displayCount);

const external = users.filter(user => user.IsShareByEmailGuestUser === true)
const externalfiltered = usersFilter === "" ? external :
      external.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
      )
const externalToDisplay = externalfiltered.slice(0, displayCount);

const handleScroll = (event) => {
    event.currentTarget.scrollTop + event.currentTarget.offsetHeight >= event.currentTarget.scrollHeight ? setDisplayCount(displayCount + 30) : null
}

//console.log(internalToDisplay)

//RESIZE
const [initialPos, setInitialPos] = useState(null);
const [initialSize, setInitialSize] = useState(null);
//const [initialPosBottom, setInitialPosBottom] = useState(null);
//const [initialSizeBottom, setInitialSizeBottom] = useState(null);

console.log(initialPos) //689
console.log(initialSize) //443
//console.log(initialPosBottom) //560
//console.log(initialSizeBottom) //684

const initial = (e) => {
    let resizable = document.getElementById('Resizable'); 
    //let resizableBottom = document.getElementById('ResizableBottom'); 
    setInitialPos(e.clientY);
    setInitialSize(resizable.offsetHeight);   
    //setInitialPosBottom(e.clientY);
    //setInitialSizeBottom(resizableBottom.offsetHeight); 
}

const resize = (e) => {
    
    console.log(e.clientY)
    let resizable = document.getElementById('Resizable'); 
    //let resizableBottom = document.getElementById('ResizableBottom');
    resizable.style.height =  `${parseInt(initialSize) + parseInt(`${e.clientY - initialPos}`)}px` 
    //resizableBottom.style.height =  `${parseInt(initialSizeBottom) - parseInt(`${e.clientY - initialPosBottom}`)}px` 
    //console.log(resizable.style.height)
    //console.log(resizableBottom.style.height)
}

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
                /><span>({displayCount}/{internalToDisplay.length})</span>
        </div>
        <div className={perstyles.permissionModalMiddle} id = 'Resizable'>
            <div className={perstyles.groupsWrapper}>
                <div className={perstyles.groupBox}>
                    <span className={perstyles.groupBoxHeader}>Admins</span>
                        <div className={perstyles.groupBoxResults}>
                                <ul>
                                    {adminfiltered.map((user,idx) => 
                                        <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
                                    )}
                                </ul>                  
                        </div>
                </div>
                <div className={perstyles.groupBox}>
                    <span className={perstyles.groupBoxHeader}>Owners</span>
                    <div className={perstyles.groupBoxResults}>
                        <ul>
                            {ownersfiltered.map((user,idx) => 
                                <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
                            )}
                        </ul>                    
                    </div>

                </div>
            </div>
            <div className={perstyles.groupsWrapper}>
                <div className={perstyles.groupBox}>
                    <span className={perstyles.groupBoxHeader}>Members</span>
                    <div className={perstyles.groupBoxResults}>
                        <ul>
                            {membersfiltered.map((user,idx) => 
                                <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
                            )}
                        </ul>
                    </div>

                </div>
                <div className={perstyles.groupBox}>
                    <span className={perstyles.groupBoxHeader}>Visitors</span>
                    <div className={perstyles.groupBoxResults}>
                        <ul >
                        {visitorsfiltered.map((user,idx) => 
                            <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
                        )}
                        </ul>
                    </div>

                </div>
            </div>
        </div>

        <div className={perstyles.draggable} id = 'Draggable'
                draggable   = 'true'
                onDragStart = {initial} 
                onDrag      = {resize}
            />


        <div className={perstyles.permissionModalBottom} id = 'ResizableBottom'>
            <div className={perstyles.groupsWrapper} onScroll={handleScroll}>
                <div className={perstyles.groupBox}>
                    <span className={perstyles.groupBoxHeader}>MSF users ({internal.length})</span>
                    <div className={perstyles.groupBoxResults}> 
                        <ul>
                            {internalToDisplay.map((user,idx) =>
                            <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
                            )}
                        </ul>
                    </div>
                </div>
            </div>
            <div className={perstyles.groupsWrapper} onScroll={handleScroll}>
                <div className={perstyles.groupBox}>
                    <span className={perstyles.groupBoxHeader}>Non-MSF users ({external.length})</span>
                    <div className={perstyles.groupBoxResults}> 
                        <ul>
                            {externalToDisplay.map((user,idx) =>
                            <PersonPermissions key={idx} user={user} context={props.context} url={props.url}/>
                            )}
                        </ul>
                    </div>
                </div>
            </div>
        </div>

    </div>
        
    );
  }


  function PersonPermissions (props) {
    const user = props.user
    
    const [up,setUp] = useState(null)
    async function getUp(loginName) {  
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const user = await site.getUserEffectivePermissions(loginName) 
        
        if (user.Low === 4294967295) {return "Full control"}
        else if (user.Low === 1011030767) {return "Edit"}
        else if (user.Low === 138612833) {return "Read"}
        else if (user.Low === 0) {return "No access"}

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


    return (
        <li>    
            <div className={perstyles.userPermBox}>
                <div>
                    <span className={user.Title === "Everyone except external users" || user.Title === "Everyone" ? perstyles.permissionWarning : ""}>
                        {user.Title} {user.Title === "Everyone except external users" || user.Title === "Everyone" ? <Icon iconName="WarningSolid"/> : null}
                    </span>
                    <span className={perstyles.userPermMail}>
                        {user.LoginName}
                    </span> 
                </div>
                <span>
                    {up}
                </span>  
            </div>        
        </li>
    )
  }


   

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

          /*
        if(
            //FULL
            sp.web.hasPermissions(user,PermissionKind.FullMask)        
        ) {
          
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
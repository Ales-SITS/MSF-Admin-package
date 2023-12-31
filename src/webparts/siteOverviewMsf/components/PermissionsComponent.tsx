import * as React from 'react';
import {useState, useEffect} from 'react';

//import styles from './SiteOverviewMsf.module.scss';
import '../CSS/styles.css';

import perstyles from './PermissionsComponent.module.scss';

import { spfi, SPFx as SPFxsp} from "@pnp/sp";

//VISUAL
//import { Resizable,ResizableBox } from 'react-resizable';

//PNP/SP
import { Web } from "@pnp/sp/webs";   
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/security";

//GRAPH
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";
import "@pnp/graph/users";
import "@pnp/graph/sites";

import { Icon } from '@fluentui/react/lib/Icon';


//FUNCTIONS
async function getGroups(web,url):Promise<any> {
    const site = Web([web, `${url}`]) 
    const groups = await site.siteGroups()
    return groups
 }

async function getGroupUsers(context,id):Promise<any>{
    const graph = graphfi().using(SPFx(context))

    const owners = await graph.groups.getById(id).owners()
    const members = await graph.groups.getById(id).members()

    return [owners,members]
}

async function getOwners(web,url):Promise<any> {
    const site = Web([web, `${url}`]) 
    const groupid = await site.associatedOwnerGroup().then(result => {return result.Id})
    const owners = await site.siteGroups.getById(groupid).users();
    return owners
 }

 async function getMembers(web,url):Promise<any> {
    const site = Web([web, `${url}`]) 
    const groupid = await site.associatedMemberGroup().then(result => {return result.Id})
    const owners = await site.siteGroups.getById(groupid).users();
    return owners
 }

 async function getVisitors(web,url):Promise<any> {
    const site = Web([web, `${url}`]) 
    const groupid = await site.associatedVisitorGroup().then(result => {return result.Id})
    const owners = await site.siteGroups.getById(groupid).users();
    return owners
 }

//COMPONENT
export default function PermissionsComponent (props) {
    const sp = props.sp

    const [topSize,setTopSize] = useState(300)
    const onResize = (event, {node, size, handle}) => {
        setTopSize(size.height);
      };

    const [everyone, setEveryone] = useState(false)
    const [users,setUsers] = useState([])
    async function getUsers() {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.siteUsers()    

        return users
    }
    
    const [userGroup,setUserGroup] = useState([])
    async function getUserGroup(id):Promise<any> {   
        const sp = spfi().using(SPFxsp(props.context));     
        const site = Web([sp.web, `${props.url}`])      
        const users = await site.siteUsers.getById(id).groups()    
    
        return users
    } 

    const [owners,setOwners] = useState([])
    const [members,setMembers] = useState([])
    const [visitors,setVisitors] = useState([])

    const [role,setRole] = useState(null)
    async function getRole():Promise<any> {   
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

        getGroupUsers(props.context,"314fa36a-0777-469f-84f3-3efb9bb8508c").then(result => console.log(result))
        //getGroups(sp.web,props.url).then(result => console.log(result.filter(group => !group.LoginName.includes("SharingLinks"))))
        //getRole()

        }, []);


//VISUAL
const closeHandler = ():void => {
    props.onCloseHandler()
}


//FILTERED
const [displayCountAll, setDisplayCountAll] = useState(25);
const [displayCountExternal, setDisplayCountExternal] = useState(25);

const [usersFilter,setUsersFilter] = useState("")
const usersFilterHandler = (e):void => {
    setUsersFilter(e)
}

const handleScroll = (event,scroll):void => {
    //console.log(event)
    if(scroll===1) event.currentTarget.scrollTop + event.currentTarget.offsetHeight >= event.currentTarget.scrollHeight ? setDisplayCountAll(displayCountAll + 25) : null
    if(scroll===2) event.currentTarget.scrollTop + event.currentTarget.offsetHeight >= event.currentTarget.scrollHeight ? setDisplayCountExternal(displayCountExternal + 25) : null
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
const internalToDisplay = internalfiltered.slice(0, displayCountAll);

const external = users.filter(user => user.IsShareByEmailGuestUser === true)
const externalfiltered = usersFilter === "" ? external :
      external.filter( user => user.Title.toLowerCase().includes(usersFilter.toLowerCase())     
      )
const externalToDisplay = externalfiltered.slice(0, displayCountExternal);

const[selected,setSelected] = useState(1)
const selectedHandler = (val) => {
    setSelected(val)
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
        <input 
                className={perstyles.permissionModalInput}
                type="text" 
                name="siteName" 
                placeholder="Filter by user name"
                onChange={e => usersFilterHandler(e.target.value)} 
                />
        <div className={perstyles.permissionModalNav}>
            <button type="button" onClick={()=>selectedHandler(1)} className={selected === 1 ? perstyles.selectedButton : null}>Sigle user search</button>
            <button type="button" onClick={()=>selectedHandler(2)} className={selected === 2 ? perstyles.selectedButton : null}>Top groups</button>
            <button type="button" onClick={()=>selectedHandler(3)} className={selected === 3 ? perstyles.selectedButton : null}>All users</button>
        </div>
        {selected === 1 && 
        <div className={perstyles.permissionModalInside}>
            <iframe className={perstyles.checkPermissionIframe} src={`${props.url}/_layouts/15/chkperm.aspx?obj=https%3a%2f%2fmsfintl.sharepoint.com%2fsites%2fmsfintlcommunities%2cWEB&IsDlg=1`}></iframe>
        </div>
        }
        {selected === 2 && 
        <div className={perstyles.permissionModalInside}>
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
        }
        {selected === 3 &&
        <div className={perstyles.permissionModalInside}>
                <div className={perstyles.groupsWrapper}>
                    <div className={perstyles.groupBox}>
                        <span className={perstyles.groupBoxHeader}>MSF users ({displayCountAll}/{internal.length})</span>
                        <div className={perstyles.groupBoxResults} onScroll={(e)=>handleScroll(e,1)}> 
                            <ul>
                                {internalToDisplay.map((user,idx) =>
                                <PersonPermissions key={idx} user={user} context={props.context} url={props.url} group={"all"}/>
                                )}
                            </ul>
                        </div>
                    </div>
                </div>
                <div className={perstyles.groupsWrapper}>
                    <div className={perstyles.groupBox}>
                        <span className={perstyles.groupBoxHeader}>Non-MSF users ({displayCountExternal}/{external.length})</span>
                        <div className={perstyles.groupBoxResults} onScroll={(e)=>handleScroll(e,2)}> 
                            <ul>
                                {externalToDisplay.map((user,idx) =>
                                <PersonPermissions key={idx} user={user} context={props.context} url={props.url} group={"external"}/>
                                )}
                            </ul>
                        </div>
                    </div>
                </div>
        </div>
        }
    </div>
        
    );
  }


  function PersonPermissions (props) {
    const user = props.user
    
    const [up,setUp] = useState(null)
    const [owners, setOwners] = useState([])
    const [members, setMembers] = useState([])

    const M365 = props.user.LoginName.startsWith("c:0o.c")&&!props.user.LoginName.endsWith("_o")&&!props.user.Email.endsWith("@msf.org")
    const M365o = props.user.LoginName.startsWith("c:0o.c")&&props.user.LoginName.endsWith("_o")&&!props.user.Email.endsWith("@msf.org")

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

        if(M365||M365o) {
            //console.log(props.user)
            const groupID = props.user.LoginName.split("|")[2].replace("_o","") 
            getGroupUsers(props.context,groupID).then(result => {
                setOwners(result[0])
                setMembers(result[1])
            })
        }
        },[props]);

        const[M365visible,setM365visible] = useState(false)

  //console.log(members)

    return (
        <li className={perstyles.userPerm}>    
            <div className={perstyles.userPermBox}>
                <div>
                    <span className={user.Title === "Everyone except external users" || user.Title === "Everyone" ? perstyles.permissionWarning : ""}>
                        {user.Title} {M365o}{user.Title === "Everyone except external users" || user.Title === "Everyone" ? <Icon iconName="WarningSolid"/> : null}
                    </span>
                    <span className={perstyles.userPermMail}>
                        {user.LoginName}
                    </span> 
                </div>
                <div className={perstyles.userPerm}>
                    <span>
                        {up}
                    </span> 
                    {M365===true || M365o === true ? 
                    <button onClick={()=>setM365visible(!M365visible)}>M365 group</button>
                    : null}
                </div> 
            </div>
            {M365o&&M365visible&&
                <ul>
                    {owners.map(user => <li>{user.displayName}</li>)}
                </ul>
            }
            {M365&&M365visible&& 
                <ul>
                    {members.map(user => <li>{user.displayName}</li>)}
                </ul>
         }     
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
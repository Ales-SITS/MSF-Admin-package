import * as React from 'react';
import {useState, useEffect, useRef} from 'react';

import scriptstyles from './PnPScripts.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';
import SyntaxHighlighter from 'react-syntax-highlighter';

import {PnPCollectionAdmin,PnPCollectionAdmin2} from './PnP_Scripts'
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';


export default function PnP_AllPermissions (props) {

    const closeHandler = () => {
        props.onCloseHandler()
    }

    const [path,setPath] = useState("")

    const scripts = [
        {key:"script1", text: "Get Admin"},
        {key:"script2", text: "Full permissions report"}
    ]
    
    const [selectedScript, setSelectedScript] = useState<IDropdownOption>(scripts[0])

    const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
    const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        setSelectedScript(item);
      };

    console.log(selectedScript)

    return (
        <div className={scriptstyles.scriptModal}>
            <div className={scriptstyles.scriptModalTop}>
                <h3>PnP Script generator</h3>
                <button onClick={closeHandler}><Icon iconName="ChromeClose"/></button>
            </div>
            <div className={scriptstyles.scriptModalInput}>
                <Dropdown
                    selectedKey={selectedScript ?selectedScript.key : undefined}
                    onChange={onChange}
                    placeholder="Select an option"
                    options={scripts}
                    styles={dropdownStyles}
                    />
                <input 
                    type="text" 
                    name="localPath" 
                    placeholder="Local path and file name, i.e: C:\Users\YourName\Desktop\Report.csv"
                    onChange={e => setPath(e.target.value)} 
                    />
            </div>
            { 
            selectedScript.key === "script1" ? <PnPCollectionAdmin siteurl={props.siteurl} reportpath={path}/> :
            selectedScript.key === "script2" ? <PnPCollectionAdmin2 siteurl={props.siteurl} reportpath={path}/> :
            null
            }
        </div>

    );
}



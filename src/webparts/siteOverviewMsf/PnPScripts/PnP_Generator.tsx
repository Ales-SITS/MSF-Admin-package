import * as React from 'react';
import {useState, useCallback} from 'react';

import scriptstyles from './PnPScripts.module.scss';

import { Icon } from '@fluentui/react/lib/Icon';


import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';

import {PnPCollectionAdmin, PnPAllPermissions} from './PnP_Scripts'


export default function PnP_AllPermissions (props) {

  const closeHandler = () => {
        props.onCloseHandler()
    }

  const topSiteScripts = [
        {key:"script1", text: "Get Admin"},
        {key:"script2", text: "Full permissions report"}
    ]
    
  const libraryScripts = [
      {key:"script3", text: "XXX"},
      {key:"script4", text: "YYY"}
  ]

  const scripts = props.type === "top_site" ? topSiteScripts : libraryScripts


    const [selectedScript, setSelectedScript] = useState<IDropdownOption>(scripts[0])
    const [inputValue, setInputValue] = useState('');

    const dropdownStyles: Partial<IDropdownStyles> = { 
        dropdown: { 
            width: 300,
            borderRadius: '5px',
            borderColor: 'orange',
            selectors: {
                ':hover': {
                  borderColor: 'orange', // Change the border color on hover
                },
                ':focus::after': {
                  borderColor: 'orange', // Change the border color on focus
                  borderRadius: '5px',
                }
         }
        },
        title: {
            borderRadius: '5px',
        }
        };

    const textfieldStylesDisabled: Partial<ITextFieldStyles> = {
            root: { 
               width: 550
            },
            field: {
               width: 550,
               borderRadius: '5px',         
            }
            ,
            fieldGroup: {
               width: 550,
               borderRadius: '5px',
              }
           };

    const textfieldStyles: Partial<ITextFieldStyles> = {
         root: { 
            width: 500
         },
         field: {
            width: 500,
            borderRadius: '5px',
            selectors: {
                ':hover': {
                  borderColor: 'orange', // Change the border color on hover
                },
                ':focus': {
                  borderColor: 'orange', // Change the border color on focus
                },
                '::after': {
                  borderColor: 'orange', // Change the border color on focus
                    borderRadius: '5px',
                }
         }
         },
         fieldGroup: {
            width: 500,
            borderRadius: '5px',
            selectors: {
                ':hover': {
                  borderColor: 'orange', // Change the border color on hover
                },
                ':focus': {
                  borderColor: 'orange', // Change the border color on focus
                },
                '::after': {
                  borderColor: 'orange', // Change the border color on focus
                  borderRadius: '5px',
                }
         }}
        };

    const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        setSelectedScript(item);
      };

    const inputValueHandler = useCallback(
        (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
          setInputValue(newValue || '');
        },
        [],
      );


    console.log(selectedScript)

    return (
        <div className={scriptstyles.scriptModal}>
            <div className={scriptstyles.scriptModalTop}>
                <h2>PnP Script generator</h2>
                <button onClick={closeHandler}><Icon iconName="ChromeClose"/></button>
            </div>
            <div className={scriptstyles.scriptModalInput}>
                <Dropdown
                    label="Select script" 
                    selectedKey={selectedScript ?selectedScript.key : undefined}
                    onChange={onChange}
                    placeholder="Select an option"
                    options={scripts}
                    styles={dropdownStyles}
                    />
                { 
                     selectedScript.key === "script1" ? 
                     <TextField 
                     label="No Additional Input needed " 
                     disabled
                     styles={textfieldStylesDisabled}
                       /> :
                     <TextField 
                     label="Additional Input " 
                     onChange={inputValueHandler}
                     placeholder={selectedScript.key === "script2" ? "Local path and file name, i.e: C:/Users/YourName/Desktop/Report.csv" : "?"}
                     required
                     styles={textfieldStyles}
                     />
                     
                }
            </div>
            { 
            selectedScript.key === "script1" ? <PnPCollectionAdmin siteurl={props.siteurl} reportpath={inputValue}/> :
            selectedScript.key === "script2" ? <PnPAllPermissions siteurl={props.siteurl} reportpath={inputValue}/> :
            null
            }
        </div>

    );
}


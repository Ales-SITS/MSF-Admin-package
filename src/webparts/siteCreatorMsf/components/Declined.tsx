import * as React from 'react';
import styles from './SiteCreatorMsf.module.scss';


export default function Declined (props) { 

  const user = props.context.pageContext.user.email.toLowerCase()

  return (
    <div className={styles.decline_wrapper}>
      {props.loader ? 
      <span className={styles.decline_loading}>Checking your permissions ...</span> :
      <span className={styles.decline_msg}>Your account {user} doesn't have permissions to create a site.</span>}
    </div>
  )
}

import * as React from 'react';
import styles from './SiteCreatorMsf.module.scss';


export default function Loader () { 

  return (
      <div className={styles.lds_ellipsis}><div></div><div></div><div></div><div></div></div>
  )
}

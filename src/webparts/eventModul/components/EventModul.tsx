import * as React from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
//import { PrimaryButton, ActionButton, IIconProps } from "@fluentui/react";

interface IEventModulState {
  isTrue: boolean;
}

export default class EventModul extends React.Component<IEventModulProps, IEventModulState> {
  
  public render(): React.ReactElement<IEventModulProps> {
    const { hasTeamsContext, userDisplayName } = this.props;


    return (
      <section
        className={`${styles.eventModul} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <h2>{escape(userDisplayName)}s Events</h2>
          <p>Her kan du se alle dine events og fremtidige events</p>
        </div>

        
      </section>
    );
  }
}

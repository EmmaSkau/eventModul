import * as React from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { PrimaryButton, ActionButton, IIconProps } from "@fluentui/react";
import CreateEvent from "./CreateEvent";

interface IEventModulState {
  showCreateEvent: boolean;
}

export default class EventModul extends React.Component<IEventModulProps, IEventModulState> {
  constructor(props: IEventModulProps) {
    super(props);
    this.state = {
      showCreateEvent: false
    };
  }

  private handleCreateEvent = (): void => {
    this.setState({ showCreateEvent: true });
  };

  private handleCloseCreateEvent = (): void => {
    this.setState({ showCreateEvent: false });
  };
  
  public render(): React.ReactElement<IEventModulProps> {
    const { hasTeamsContext, userDisplayName } = this.props;

    const monthIcon: IIconProps = { iconName: "Calendar" };
    const thisYearIcon: IIconProps = { iconName: "CalendarYear" };

    return (
      <section
        className={`${styles.eventModul} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <h2>{escape(userDisplayName)}s Events</h2>
          <p>Her kan du se alle dine events og oprette ny events</p>
        </div>

        <PrimaryButton 
          text="Opret ny event" 
          onClick={this.handleCreateEvent}
        />

        {this.state.showCreateEvent && (
          <CreateEvent onClose={this.handleCloseCreateEvent} />
        )}

        <div>
          <ActionButton
            iconProps={monthIcon}
            allowDisabledFocus
            disabled={false}
            checked={false}
          >
            Denne måned
          </ActionButton>
          <ActionButton
            iconProps={thisYearIcon}
            allowDisabledFocus
            disabled={false}
            checked={false}
          >
            Dette år
          </ActionButton>
        </div>
      </section>
    );
  }
}

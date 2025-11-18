import * as React from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { PrimaryButton, ActionButton, IIconProps } from "@fluentui/react";
import CreateEvent from "./CreateEvent";
import AdminListView from "./ListView/AdminListView";

interface IAdminPageProps extends IEventModulProps {
  onClose?: () => void;
}

interface IEventModulState {
  showCreateEvent: boolean;
}

export default class AdminPage extends React.Component<
  IAdminPageProps,
  IEventModulState
> {
  private listViewRef = React.createRef<AdminListView>();

  constructor(props: IAdminPageProps) {
    super(props);
    this.state = {
      showCreateEvent: false,
    };
  }

  private handleCreateEvent = (): void => {
    this.setState({ showCreateEvent: true });
  };

  private handleCloseCreateEvent = (): void => {
    this.setState({ showCreateEvent: false });
  };

  private handleEventCreated = (): void => {
    // Refresh the ListView when a new event is created
    if (this.listViewRef.current) {
      // Try multiple refreshes with increasing delays to catch the new item once SharePoint indexes it
      const refreshTimes = [2000, 4000, 6000]; // Refresh at 2s, 4s, and 6s

      refreshTimes.forEach((delay) => {
        setTimeout(() => {
          if (this.listViewRef.current) {
            this.listViewRef.current.loadEvents().catch(console.error);
          }
        }, delay);
      });
    }
  };

  private handleCloseAdminPage = (): void => {
    if (this.props.onClose) {
      this.props.onClose();
    }
  };

  public render(): React.ReactElement<IAdminPageProps> {
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
          <PrimaryButton
            text="Luk Admin page"
            onClick={this.handleCloseAdminPage}
          />
        </div>

        <PrimaryButton text="Opret ny event" onClick={this.handleCreateEvent} />

        <CreateEvent
          isOpen={this.state.showCreateEvent}
          onClose={this.handleCloseCreateEvent}
          context={this.props.context}
          onEventCreated={this.handleEventCreated}
        />

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
        <AdminListView ref={this.listViewRef} context={this.props.context} />
      </section>
    );
  }
}

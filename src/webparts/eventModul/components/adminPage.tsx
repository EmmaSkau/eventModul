import * as React from "react";
import { useState, useCallback } from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./Utility/IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { PrimaryButton, ActionButton, IIconProps } from "@fluentui/react";
import CreateEvent from "./CreateEvent";
import AdminListView from "./ListView/AdminListView";

interface IAdminPageProps extends IEventModulProps {
  onClose?: () => void;
}

const AdminPage: React.FC<IAdminPageProps> = (props) => {
  // State
  const [showCreateEvent, setShowCreateEvent] = useState(false);
  const [refreshTrigger, setRefreshTrigger] = useState(0);

  // Event handlers
  const handleCreateEvent = useCallback((): void => {
    setShowCreateEvent(true);
  }, []);

  const handleCloseCreateEvent = useCallback((): void => {
    setShowCreateEvent(false);
  }, []);

  const handleEventCreated = useCallback((): void => {
    setRefreshTrigger((prev) => prev + 1);

    const refreshTimes = [2000, 4000, 6000]; // Refresh at 2s, 4s, and 6s
    refreshTimes.forEach((delay) => {
      setTimeout(() => {
        setRefreshTrigger((prev) => prev + 1);
      }, delay);
    });
  }, []);

  const handleCloseAdminPage = useCallback((): void => {
    if (props.onClose) {
      props.onClose();
    }
  }, [props.onClose]);

  // Render
  const { hasTeamsContext, userDisplayName } = props;

  const monthIcon: IIconProps = { iconName: "Calendar" };
  const thisYearIcon: IIconProps = { iconName: "CalendarYear" };

  return (
    <section
      className={`${styles.eventModul} ${hasTeamsContext ? styles.teams : ""}`}
    >
      <div className={styles.welcome}>
        <h2>{escape(userDisplayName)}s Events</h2>
        <p>Her kan du se alle dine events og oprette ny events</p>
        <PrimaryButton text="Luk Admin page" onClick={handleCloseAdminPage} />
      </div>

      <PrimaryButton text="Opret ny event" onClick={handleCreateEvent} />

      <CreateEvent
        isOpen={showCreateEvent}
        onClose={handleCloseCreateEvent}
        context={props.context}
        onEventCreated={handleEventCreated}
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
      <AdminListView key={refreshTrigger} context={props.context} />
    </section>
  );
};

export default AdminPage;

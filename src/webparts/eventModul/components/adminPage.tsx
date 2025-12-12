import * as React from "react";
import { useState, useCallback,  } from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./Utility/IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { PrimaryButton, Toggle} from "@fluentui/react";
import CreateEvent from "./CreateEvent";
import AdminListView from "./ListView/AdminListView";
import FilterUtility, { useLocationFilter } from "./Utility/filterUtility";

interface IAdminPageProps extends IEventModulProps {
  onClose?: () => void;
}

const AdminPage: React.FC<IAdminPageProps> = (props) => {
  // Use the filter hook to get filter state
  const {
    startDate,
    endDate,
    selectedLocation,
    locationOptions,
    isLoadingLocations,
    onLocationChange,
    onStartDateSelect,
    onEndDateSelect,
    resetFilters,
  } = useLocationFilter(props.context);

  // State
  const [showCreateEvent, setShowCreateEvent] = useState(false);
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [showPastEvents, setShowPastEvents] = useState(false);

  // Event handlers
  const handleCreateEvent = useCallback((): void => {
    setShowCreateEvent(true);
  }, []);

  const handleCloseCreateEvent = useCallback((): void => {
    setShowCreateEvent(false);
  }, []);

  const handleEventCreated = useCallback((): void => {
    setRefreshTrigger((prev) => prev + 1);
  }, []);

  const handleCloseAdminPage = useCallback((): void => {
    if (props.onClose) {
      props.onClose();
    }
  }, [props.onClose]);

  // Render
  const { hasTeamsContext, userDisplayName } = props;

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

      <FilterUtility 
        context={props.context}
        startDate={startDate}
        endDate={endDate}
        selectedLocation={selectedLocation}
        locationOptions={locationOptions}
        isLoadingLocations={isLoadingLocations}
        onLocationChange={onLocationChange}
        onStartDateSelect={onStartDateSelect}
        onEndDateSelect={onEndDateSelect}
        resetFilters={resetFilters}
      />
      
      <Toggle
        label="Afholdte events"
        checked={showPastEvents}
        onChange={(_, checked) => setShowPastEvents(!!checked)}
      />

      <CreateEvent
        isOpen={showCreateEvent}
        onClose={handleCloseCreateEvent}
        context={props.context}
        onEventCreated={handleEventCreated}
      />

      <AdminListView 
        key={refreshTrigger} 
        context={props.context} 
        startDate={startDate}
        endDate={endDate}
        selectedLocation={selectedLocation}
        showPastEvents={showPastEvents}
      />
    </section>
  );
};

export default AdminPage;

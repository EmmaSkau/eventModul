import * as React from "react";
import { useState, useCallback } from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./Utility/IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  Toggle,
  DefaultButton,
  PrimaryButton,
} from "@fluentui/react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import ListView from "./ListView/ListView";
import RegisteredListView from "../components/ListView/RegisteredListView";
import AdminPage from "./adminPage";
import FilterUtility, { useLocationFilter } from "./Utility/filterUtility";


const EventModul: React.FC<IEventModulProps> = (props) => {
  const { hasTeamsContext, userDisplayName, context } = props;
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
  const [registered, setRegistered] = useState(false);
  const [waitlisted, setWaitlisted] = useState(false);
  const [showAdminPage, setShowAdminPage] = useState(false);
  const [refreshTrigger, setRefreshTrigger] = useState(0);

  const handleRefresh = useCallback((): void => {
    setRefreshTrigger((prev) => prev + 1);
  }, []);

  const handleOpenAdminPage = useCallback((): void => {
    setShowAdminPage(true);
  }, []);

  const handleCloseAdminPage = useCallback((): void => {
    setShowAdminPage(false);
  }, []);

  return (
    <section
      className={`${styles.eventModul} ${hasTeamsContext ? styles.teams : ""}`}
    >
      {showAdminPage && <AdminPage {...props} onClose={handleCloseAdminPage} />}

      <div className={styles.welcome}>
        <h2>{escape(userDisplayName)}s Events</h2>
        <p>Her kan du se alle dine events og fremtidige events</p>
        <PrimaryButton text="Admin page" onClick={handleOpenAdminPage} />
      </div>

      
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

      <section className={styles.filterToggle}>
        <Toggle
          label="Tilmeldt"
          checked={registered}
          onChange={(_, checked) => {
            setRegistered(!!checked);
            if (checked) {
              setWaitlisted(false);
            }
          }}
        />

        <Toggle
          label="Venteliste"
          checked={waitlisted}
          onChange={(_, checked) => {
            setWaitlisted(!!checked);
            if (checked) {
              setRegistered(false);
            }
          }}
        />

        <DefaultButton
          text="Opdater liste"
          iconProps={{ iconName: "Refresh" }}
          onClick={handleRefresh}
        />
      </section>

      <h2>
        {registered || waitlisted ? "Mine events:" : "Fremtidige events:"}
      </h2>
      {registered || waitlisted ? (
        <RegisteredListView
          key={`${refreshTrigger}-${waitlisted}`}
          context={context}
          startDate={startDate}
          endDate={endDate}
          selectedLocation={selectedLocation}
          registered={registered}
          waitlisted={waitlisted}
        />
      ) : (
        <ListView
          key={refreshTrigger}
          context={context}
          startDate={startDate}
          endDate={endDate}
          selectedLocation={selectedLocation}
        />
      )}
    </section>
  );
};

export default EventModul;

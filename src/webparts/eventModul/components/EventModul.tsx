import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./Utility/IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  DatePicker,
  DayOfWeek,
  Dropdown,
  IDropdownOption,
  Toggle,
  DefaultButton,
  PrimaryButton,
} from "@fluentui/react";
import { formatDate } from "./Utility/formatDate";
import { getSP } from "../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import ListView from "./ListView/ListView";
import RegisteredListView from "../components/ListView/RegisteredListView";
import AdminPage from "./adminPage";

interface ILocationItem {
  Placering: string;
}

const EventModul: React.FC<IEventModulProps> = (props) => {
  const { hasTeamsContext, userDisplayName, context } = props;

  // State
  const [startDate, setStartDate] = useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = useState<Date | undefined>(undefined);
  const [selectedLocation, setSelectedLocation] = useState<string | undefined>(undefined);
  const [locationOptions, setLocationOptions] = useState<IDropdownOption[]>([]);
  const [isLoadingLocations, setIsLoadingLocations] = useState(false);
  const [registered, setRegistered] = useState(false);
  const [cancelledEvents, setCancelledEvents] = useState(false);
  const [waitlisted, setWaitlisted] = useState(false);
  const [showAdminPage, setShowAdminPage] = useState(false);
  const [refreshTrigger, setRefreshTrigger] = useState(0);

  const loadLocationsFromSharePoint = useCallback(async (): Promise<void> => {
    try {
      setIsLoadingLocations(true);
      const sp = getSP(context);

      // Get all items with Placering field (Location field can't be filtered in query)
      const items: ILocationItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.select("Placering")();

      console.log("Loaded items from SharePoint:", items);

      const locations = items
        .map((item: ILocationItem) => {
          if (!item.Placering) return null;
          try {
            const parsed = JSON.parse(item.Placering);
            return parsed.DisplayName || item.Placering;
          } catch {
            return item.Placering;
          }
        })
        .filter((location) => location);
      const uniqueLocations: string[] = [];
      const seen: { [key: string]: boolean } = {};
      for (const location of locations) {
        if (!seen[location]) {
          seen[location] = true;
          uniqueLocations.push(location);
        }
      }

      const options: IDropdownOption[] = [
        { key: "all", text: "Alle lokationer" },
        ...uniqueLocations.map((location: string) => ({
          key: location,
          text: location,
        })),
      ];

      setLocationOptions(options);
      setIsLoadingLocations(false);
    } catch (error) {
      console.error("Error loading locations from SharePoint:", error);
      setIsLoadingLocations(false);
      setLocationOptions([{ key: "all", text: "Alle lokationer" }]);
    }
  }, [context]);

  // Load locations on mount
  useEffect(() => {
    loadLocationsFromSharePoint().catch((error) => {
      console.error("Error in componentDidMount:", error);
    });
  }, [loadLocationsFromSharePoint]);

  const onLocationChange = useCallback((
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      setSelectedLocation(option.key as string);
    }
  }, []);

  const onStartDateSelect = useCallback((date: Date | null | undefined): void => {
    setStartDate(date || undefined);
  }, []);

  const onEndDateSelect = useCallback((date: Date | null | undefined): void => {
    setEndDate(date || undefined);
  }, []);

  const resetFilters = useCallback((): void => {
    setStartDate(undefined);
    setEndDate(undefined);
    setSelectedLocation(undefined);
    setRegistered(false);
    setCancelledEvents(false);
    setWaitlisted(false);
  }, []);

  const handleRefresh = useCallback((): void => {
    setRefreshTrigger(prev => prev + 1);
  }, []);

  const handleOpenAdminPage = useCallback((): void => {
    setShowAdminPage(true);
  }, []);

  const handleCloseAdminPage = useCallback((): void => {
    setShowAdminPage(false);
  }, []);

  return (
    <section
      className={`${styles.eventModul} ${
        hasTeamsContext ? styles.teams : ""
      }`}
    >
      {showAdminPage && (
        <AdminPage {...props} onClose={handleCloseAdminPage} />
      )}

      <div className={styles.welcome}>
        <h2>{escape(userDisplayName)}s Events</h2>
        <p>Her kan du se alle dine events og fremtidige events</p>
        <PrimaryButton text="Admin page" onClick={handleOpenAdminPage} />
      </div>

      <section className={styles.filters}>
        <DatePicker
          label="Fra"
          firstDayOfWeek={DayOfWeek.Monday}
          placeholder="Vælg start dato..."
          ariaLabel="Vælg start dato"
          value={startDate}
          onSelectDate={onStartDateSelect}
          formatDate={formatDate}
        />

        <DatePicker
          label="Til"
          firstDayOfWeek={DayOfWeek.Monday}
          placeholder="Vælg slut dato..."
          ariaLabel="Vælg slut dato"
          value={endDate}
          onSelectDate={onEndDateSelect}
          formatDate={formatDate}
        />

        <Dropdown
          label="Lokation"
          placeholder={
            isLoadingLocations ? "Indlæser lokationer..." : "Vælg lokation..."
          }
          options={locationOptions}
          selectedKey={selectedLocation}
          onChange={onLocationChange}
          disabled={isLoadingLocations}
        />
      </section>
      <section className={styles.filterToggle}>
        <Toggle
          label="Tilmeldt"
          checked={registered}
          onChange={(_, checked) => setRegistered(!!checked)}
        />

        <Toggle label="Afmeldt" checked={cancelledEvents} />

        <Toggle label="Venteliste" checked={waitlisted} />

        <DefaultButton
          className={styles.restFilter}
          text="Ryd filtre"
          onClick={resetFilters}
        />
        <DefaultButton
          text="Opdater liste"
          iconProps={{ iconName: "Refresh" }}
          onClick={handleRefresh}
        />
      </section>

      <h2>{registered ? "Mine events:" : "Fremtidige events:"}</h2>
      {registered ? (
        <RegisteredListView
          key={refreshTrigger}
          context={context}
          startDate={startDate}
          endDate={endDate}
          selectedLocation={selectedLocation}
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

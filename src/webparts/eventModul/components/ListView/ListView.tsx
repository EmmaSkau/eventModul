import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { getSP } from "../../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  PrimaryButton,
} from "@fluentui/react";
import RegisterForEvents from "../RegisterForEvents";
import ConfirmDialog from "../Utility/ConfirmDialog";
import { IListViewProps } from "../Utility/IListViewProps";
import { IEventItem } from "../Utility/IEventItem";
import { getFutureEventsSorted, formatDate } from "../Utility/formatDate";

const ListView: React.FC<IListViewProps> = (props) => {
  // State declarations
  const [events, setEvents] = useState<IEventItem[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [registerPanelOpen, setRegisterPanelOpen] = useState(false);
  const [selectedEventId, setSelectedEventId] = useState<number | undefined>();
  const [selectedEventTitle, setSelectedEventTitle] = useState<
    string | undefined
  >();
  const [registeredEventIds, setRegisteredEventIds] = useState<number[]>([]);
  const [registrationCounts, setRegistrationCounts] = useState<{
    [eventId: number]: number;
  }>({});
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [confirmDialogConfig, setConfirmDialogConfig] = useState<{
    title: string;
    message: string;
    onConfirm: () => void;
  }>({ title: "", message: "", onConfirm: () => {} });
  const [successMessage, setSuccessMessage] = useState<string | undefined>();
  const [errorMessage, setErrorMessage] = useState<string | undefined>();

  // Load registration counts
  const loadRegistrationCounts = useCallback(async (): Promise<void> => {
    try {
      const sp = getSP(props.context);
      const timestamp = Date.now();

      // Get all registrations
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`Id ge 0 or Id eq ${timestamp}`)
        .select("EventId", "EventType")
        .top(5000)();

      // Count registrations per event
      const counts: { [eventId: number]: number } = {};

      registrations.forEach((reg) => {
        // Count if EventType is not explicitly 'Waitlist'
        if (reg.EventId && reg.EventType !== "Waitlist") {
          counts[reg.EventId] = (counts[reg.EventId] || 0) + 1;
        }
      });

      setRegistrationCounts(counts);
    } catch (error) {
      console.error("Error loading registration counts:", error);
    }
  }, [props.context]);

  // Filter events based on props
  const filterEvents = useCallback(
    (items: IEventItem[]): IEventItem[] => {
      let filtered = [...items];

      if (props.startDate) {
        filtered = filtered.filter((item) => {
          if (!item.Dato) return false;
          const eventDate = new Date(item.Dato);
          return eventDate >= props.startDate!;
        });
      }

      if (props.endDate) {
        filtered = filtered.filter((item) => {
          if (!item.Dato) return false;
          const eventDate = new Date(item.Dato);
          return eventDate <= props.endDate!;
        });
      }

      if (props.selectedLocation && props.selectedLocation !== "all") {
        filtered = filtered.filter((item) => {
          if (!item.Placering) return false;
          try {
            const parsed = JSON.parse(item.Placering);
            return parsed.DisplayName === props.selectedLocation;
          } catch {
            return item.Placering === props.selectedLocation;
          }
        });
      }

      filtered = getFutureEventsSorted(filtered);
      return filtered;
    },
    [props.startDate, props.endDate, props.selectedLocation]
  );

  // Load events
  const loadEvents = useCallback(async (): Promise<void> => {
    try {
      setIsLoading(true);
      setError(undefined);
      const sp = getSP(props.context);
      const timestamp = Date.now();

      const items: IEventItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.filter(`Id ge 0 or Id eq ${timestamp}`)
        .select(
          "Id",
          "Title",
          "Dato",
          "SlutDato",
          "Administrator/Title",
          "Placering",
          "Capacity"
        )
        .expand("Administrator")
        .top(1000)();

      const filteredItems = filterEvents(items);

      setEvents(filteredItems);
      setIsLoading(false);
    } catch (error) {
      console.error("Error loading events:", error);
      setIsLoading(false);
      setError("Kunne ikke indläse events fra SharePoint");
    }
  }, [props.context, filterEvents]);

  // Load user registrations
  const loadUserRegistrations = useCallback(async (): Promise<void> => {
    try {
      const sp = getSP(props.context);
      const currentUser = await sp.web.currentUser();
      const timestamp = Date.now();

      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(
          `Title eq '${currentUser.Title}' and (Id ge 0 or Id eq ${timestamp})`
        )
        .select("EventId", "EventType")
        .top(5000)();

      const registeredEventIds = registrations
        .map((reg) => reg.EventId)
        .filter((id) => id);

      setRegisteredEventIds(registeredEventIds);
    } catch (error) {
      console.error("Error loading user registrations:", error);
    }
  }, [props.context]);

  // Load data on mount
  useEffect(() => {
    loadEvents().catch(console.error);
    loadUserRegistrations().catch(console.error);
    loadRegistrationCounts().catch(console.error);
  }, []);

  // Reload when filters change
  useEffect(() => {
    loadEvents().catch(console.error);
  }, [loadEvents]);

  // Check if event has custom fields
  const checkIfEventHasCustomFields = useCallback(
    async (eventId: number): Promise<boolean> => {
      try {
        const sp = getSP(props.context);
        const timestamp = Date.now();
        const fields = await sp.web.lists
          .getByTitle("EventFields")
          .items.filter(
            `EventId eq ${eventId} and (Id ge 0 or Id eq ${timestamp})`
          )
          .select("Id")
          .top(5000)();

        return fields.length > 0;
      } catch (error) {
        console.error("Error checking custom fields:", error);
        return false;
      }
    },
    [props.context]
  );

  // Register user to event
  const registerUserToEvent = useCallback(
    async (eventId: number, isWaitlist: boolean = false): Promise<void> => {
      try {
        const sp = getSP(props.context);
        const currentUser = await sp.web.currentUser();

        const registrationKey = `${eventId}_${
          props.context.pageContext.user.loginName
        }_${new Date().getTime()}`;

        await sp.web.lists.getByTitle("EventRegistrations").items.add({
          Title: currentUser.Title,
          EventId: eventId,
          BrugerId: currentUser.Id,
          RegistrationKey: registrationKey,
          Submitted: new Date().toISOString(),
          EventType: isWaitlist ? "Waitlist" : "Registered",
        });

        if (isWaitlist) {
          setSuccessMessage("Du er tilføjet til ventelisten!");
        } else {
          setSuccessMessage("Du er nu tilmeldt eventet!");
        }

        await loadUserRegistrations();
        await loadEvents();
        await loadRegistrationCounts();
      } catch (error) {
        console.error("Error registering for event:", error);
        setErrorMessage("Fejl ved tilmelding. Prøv igen.");
      }
    },
    [props.context, loadUserRegistrations, loadEvents, loadRegistrationCounts]
  );

  const handleRegister = async (
    eventId: number,
    eventTitle: string,
    isWaitlist: boolean = false
  ): Promise<void> => {
    const hasCustomFields = await checkIfEventHasCustomFields(eventId);

    if (hasCustomFields) {
      setRegisterPanelOpen(true);
      setSelectedEventId(eventId);
      setSelectedEventTitle(eventTitle);
    } else {
      const message = isWaitlist
        ? `Eventet er fuldt booket. Vil du tilføjes til ventelisten?`
        : `Vil du tilmelde dig til dette event?`;

      setConfirmDialogConfig({
        title: eventTitle,
        message: message,
        onConfirm: async () => {
          setShowConfirmDialog(false);
          await registerUserToEvent(eventId, isWaitlist);
        },
      });
      setShowConfirmDialog(true);
    }
  };

  // Close register panel
  const closeRegisterPanel = (): void => {
    setRegisterPanelOpen(false);
    setSelectedEventId(undefined);
    setSelectedEventTitle(undefined);
    loadEvents().catch(console.error);
  };

  const getColumns = (): IColumn[] => {
    return [
      {
        key: "Title",
        name: "Titel",
        fieldName: "Title",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "StartDato",
        name: "Start Dato",
        fieldName: "Dato",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.Dato ? formatDate(item.Dato) : "-";
        },
      },
      {
        key: "SlutDato",
        name: "Slut Dato",
        fieldName: "SlutDato",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.SlutDato ? formatDate(item.SlutDato) : "-";
        },
      },
      {
        key: "Administrator",
        name: "Administrator",
        fieldName: "Administrator",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.Administrator?.Title || "-";
        },
      },
      {
        key: "Placering",
        name: "Placering",
        fieldName: "Placering",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IEventItem) => {
          if (!item.Placering) return "-";
          try {
            const parsed = JSON.parse(item.Placering);
            return parsed.DisplayName || item.Placering;
          } catch {
            return item.Placering;
          }
        },
      },
      {
        key: "Capacity",
        name: "Pladser",
        fieldName: "Capacity",
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: IEventItem) => {
          const count = registrationCounts[item.Id] || 0;
          const capacity = item.Capacity || 0;
          const available = capacity - count;
          return capacity > 0 ? available.toString() : "-";
        },
      },
      {
        key: "Tilmeld",
        name: "Tilmeld",
        fieldName: "Tilmeld",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IEventItem) => {
          const isRegistered = registeredEventIds.indexOf(item.Id) !== -1;
          const count = registrationCounts[item.Id] || 0;
          const capacity = item.Capacity || 0;
          const available = capacity - count;
          const isFull = capacity > 0 && available <= 0;

          return (
            <PrimaryButton
              text={
                isRegistered ? "Tilmeldt" : isFull ? "Venteliste" : "Tilmeld"
              }
              onClick={() => handleRegister(item.Id, item.Title, isFull)}
              disabled={isRegistered}
            />
          );
        },
      },
    ];
  };

  // Render
  if (isLoading) {
    return <Spinner size={SpinnerSize.large} label="Indlæser events..." />;
  }

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
    );
  }

  if (events.length === 0) {
    return (
      <MessageBar messageBarType={MessageBarType.info}>
        Ingen events fundet
      </MessageBar>
    );
  }

  return (
    <>
      {successMessage && (
        <MessageBar
          messageBarType={MessageBarType.success}
          onDismiss={() => setSuccessMessage(undefined)}
          dismissButtonAriaLabel="Luk"
        >
          {successMessage}
        </MessageBar>
      )}
      {errorMessage && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setErrorMessage(undefined)}
          dismissButtonAriaLabel="Luk"
        >
          {errorMessage}
        </MessageBar>
      )}
      <DetailsList
        items={events}
        columns={getColumns()}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.justified}
        isHeaderVisible={true}
      />

      {registerPanelOpen && selectedEventId && selectedEventTitle && (
        <RegisterForEvents
          context={props.context}
          eventId={selectedEventId}
          eventTitle={selectedEventTitle}
          isOpen={registerPanelOpen}
          onDismiss={closeRegisterPanel}
        />
      )}

      <ConfirmDialog
        hidden={!showConfirmDialog}
        title={confirmDialogConfig.title}
        message={confirmDialogConfig.message}
        confirmText="Ja, tilmeld"
        cancelText="Annuller"
        onConfirm={confirmDialogConfig.onConfirm}
        onCancel={() => setShowConfirmDialog(false)}
      />
    </>
  );
};

export default ListView;

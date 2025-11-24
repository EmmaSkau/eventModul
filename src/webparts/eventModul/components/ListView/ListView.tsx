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
  const [selectedEventTitle, setSelectedEventTitle] = useState<string | undefined>();
  const [registeredEventIds, setRegisteredEventIds] = useState<number[]>([]);

  // Filter events based on props
  const filterEvents = useCallback((items: IEventItem[]): IEventItem[] => {
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
  }, [props.startDate, props.endDate, props.selectedLocation]);

  // Load events
  const loadEvents = useCallback(async (): Promise<void> => {
    try {
      setIsLoading(true);
      setError(undefined);
      const sp = getSP(props.context);

      const items: IEventItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.select(
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

      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`Title eq '${currentUser.Title}'`)
        .select("EventId")();

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
  }, []);

  // Reload when filters change
  useEffect(() => {
    loadEvents().catch(console.error);
  }, [loadEvents]);

  // Check if event has custom fields
  const checkIfEventHasCustomFields = useCallback(async (eventId: number): Promise<boolean> => {
    try {
      const sp = getSP(props.context);
      const fields = await sp.web.lists
        .getByTitle("EventFields")
        .items.filter(`EventId eq ${eventId}`)
        .select("Id")();

      return fields.length > 0;
    } catch (error) {
      console.error("Error checking custom fields:", error);
      return false;
    }
  }, [props.context]);

  // Register user to event
  const registerUserToEvent = useCallback(async (eventId: number): Promise<void> => {
    try {
      const sp = getSP(props.context);
      const currentUser = await sp.web.currentUser();

      const registrationKey = `${eventId}_${props.context.pageContext.user.loginName}_${new Date().getTime()}`;

      await sp.web.lists.getByTitle("EventRegistrations").items.add({
        Title: currentUser.Title,
        EventId: eventId,
        BrugerId: currentUser.Id,
        RegistrationKey: registrationKey,
        Submitted: new Date().toISOString(),
      });

      alert("Du er nu tilmeldt eventet!");

      await loadUserRegistrations();
      await loadEvents();
    } catch (error) {
      console.error("Error registering for event:", error);
      alert("Fejl ved tilmelding. Prøv igen.");
    }
  }, [props.context, loadUserRegistrations, loadEvents]);

  const handleRegister = async (eventId: number, eventTitle: string): Promise<void> => {
    const hasCustomFields = await checkIfEventHasCustomFields(eventId);

    if (hasCustomFields) {
      setRegisterPanelOpen(true);
      setSelectedEventId(eventId);
      setSelectedEventTitle(eventTitle);
    } else {
      if (confirm(`Vil du tilmelde dig til "${eventTitle}"?`)) {
        await registerUserToEvent(eventId);
      }
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
        key: "targetGroup",
        name: "Målgruppe",
        fieldName: "targetGroup",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
      },
      {
        key: "Capacity",
        name: "Kapacitet",
        fieldName: "Capacity",
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
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

          return (
            <PrimaryButton
              text={isRegistered ? "Tilmeldt" : "Tilmeld"}
              onClick={() => handleRegister(item.Id, item.Title)}
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
    </>
  );
};

export default ListView;

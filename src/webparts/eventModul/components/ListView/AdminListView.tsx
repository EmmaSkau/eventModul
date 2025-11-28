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
  DefaultButton,
  ActionButton,
} from "@fluentui/react";
import CreateEvent from "../CreateEvent";
import { IListViewProps } from "../Utility/IListViewProps";
import { IEventItem } from "../Utility/IEventItem";
import { getFutureEventsSorted, formatDate } from "../Utility/formatDate";
import ManageRegistrations from "../ManageRegistrations";

const AdminListView: React.FC<IListViewProps> = (props) => {
  // State declarations
  const [events, setEvents] = useState<IEventItem[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [editPanelOpen, setEditPanelOpen] = useState(false);
  const [selectedEventForEdit, setSelectedEventForEdit] = useState<
    IEventItem | undefined
  >();
  const [registrationCounts, setRegistrationCounts] = useState<{
    [eventId: number]: number;
  }>({});
  const [manageRegistrationsOpen, setManageRegistrationsOpen] = useState(false);
  const [selectedEventForRegistrations, setSelectedEventForRegistrations] =
    useState<IEventItem | undefined>();

  // Load registration counts
  const loadRegistrationCounts = useCallback(async (): Promise<void> => {
    try {
      const sp = getSP(props.context);

      // Get all registrations
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.select("EventId")();

      // Count registrations per event
      const counts: { [eventId: number]: number } = {};

      registrations.forEach((reg) => {
        if (reg.EventId) {
          counts[reg.EventId] = (counts[reg.EventId] || 0) + 1;
        }
      });

      setRegistrationCounts(counts);
    } catch (error) {
      console.error("Error loading registration counts:", error);
    }
  }, [props.context]);

  // Filter events
  const filterEvents = useCallback(
    (items: IEventItem[]): IEventItem[] => {
      let filtered = [...items];

      // Filter by start date
      if (props.startDate) {
        filtered = filtered.filter((item) => {
          if (!item.Dato) return false;
          const eventDate = new Date(item.Dato);
          return eventDate >= props.startDate!;
        });
      }

      // Filter by end date
      if (props.endDate) {
        filtered = filtered.filter((item) => {
          if (!item.Dato) return false;
          const eventDate = new Date(item.Dato);
          return eventDate <= props.endDate!;
        });
      }

      // Filter by location
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
      const currentUserEmail = props.context.pageContext.user.email;

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
        .filter(`Administrator/EMail eq '${currentUserEmail}'`)
        .top(1000)();

      const filteredItems = filterEvents(items);

      setEvents(filteredItems);
      setIsLoading(false);
    } catch (error) {
      console.error("Error loading events:", error);
      setIsLoading(false);
      setError("Kunne ikke indlæse events fra SharePoint");
    }
  }, [props.context, filterEvents]);

  // Load data on mount
  useEffect(() => {
    loadRegistrationCounts().catch(console.error);
    loadEvents().catch(console.error);
  }, []);

  // Reload when filters change
  useEffect(() => {
    loadEvents().catch(console.error);
  }, [loadEvents]);

  // Event handlers
  const handleEditEvent = (item: IEventItem): void => {
    setEditPanelOpen(true);
    setSelectedEventForEdit(item);
  };

  const handleCloseEditPanel = (): void => {
    setEditPanelOpen(false);
    setSelectedEventForEdit(undefined);
    loadRegistrationCounts().catch(console.error);
  };

  const handleDeleteEvent = async (item: IEventItem): Promise<void> => {
    if (!confirm(`Er du sikker på, at du vil slette "${item.Title}"?`)) {
      return;
    }

    try {
      const sp = getSP(props.context);

      await sp.web.lists.getByTitle("EventDB").items.getById(item.Id).delete();

      alert("Event slettet!");

      await loadEvents();
      await loadRegistrationCounts();
    } catch (error) {
      console.error("Error deleting event:", error);
      alert("Fejl ved sletning af event. Prøv igen.");
    }
  };

  const openManageRegistrations = (item: IEventItem): void => {
    setManageRegistrationsOpen(true);
    setSelectedEventForRegistrations(item);
  };

  const closeManageRegistrations = (): void => {
    setManageRegistrationsOpen(false);
    setSelectedEventForRegistrations(undefined);
    loadRegistrationCounts().catch(console.error);
  };

  const getColumns = (): IColumn[] => {
    return [
      {
        key: "Title",
        name: "Titel",
        fieldName: "Title",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
      },
      {
        key: "StartDato",
        name: "Start Dato",
        fieldName: "Dato",
        minWidth: 60,
        maxWidth: 80,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.Dato ? formatDate(item.Dato) : "-";
        },
      },
      {
        key: "SlutDato",
        name: "Slut Dato",
        fieldName: "SlutDato",
        minWidth: 60,
        maxWidth: 80,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.SlutDato ? formatDate(item.SlutDato) : "-";
        },
      },
      {
        key: "Administrator",
        name: "Administrator",
        fieldName: "Administrator",
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.Administrator?.Title || "-";
        },
      },
      {
        key: "Placering",
        name: "Placering",
        fieldName: "Placering",
        minWidth: 80,
        maxWidth: 100,
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
        minWidth: 80,
        maxWidth: 120,
        isResizable: true,
      },
      {
        key: "tilmeldt",
        name: "Tilmeldte",
        fieldName: "tilmeldt",
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: IEventItem) => {
          const count = registrationCounts[item.Id] || 0;
          const capacity = item.Capacity || 0;
          const displayText =
            capacity > 0 ? `${count}/${capacity}` : count.toString();
          return (
            <div style={{ display: "flex", alignItems: "center" }}>
              <span>{displayText}</span>
              <ActionButton
                iconProps={{ iconName: "Edit" }}
                title="Administrer tilmeldte"
                onClick={() => openManageRegistrations(item)}
              />
            </div>
          );
        },
      },
      {
        key: "actions",
        name: "Handlinger",
        fieldName: "actions",
        minWidth: 180,
        maxWidth: 400,
        isResizable: true,
        flexGrow: 1,
        onRender: (item: IEventItem) => {
          return (
            <div style={{ display: "flex", gap: "8px" }}>
              <DefaultButton text="Ret" onClick={() => handleEditEvent(item)} />
              <PrimaryButton
                text="Slet"
                onClick={() => handleDeleteEvent(item)}
              />
            </div>
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

      <CreateEvent
        isOpen={editPanelOpen}
        onClose={handleCloseEditPanel}
        context={props.context}
        onEventCreated={loadEvents}
        eventToEdit={selectedEventForEdit}
      />

      {manageRegistrationsOpen && selectedEventForRegistrations && (
        <ManageRegistrations
          isOpen={manageRegistrationsOpen}
          onDismiss={closeManageRegistrations}
          context={props.context}
          eventId={selectedEventForRegistrations.Id}
          eventTitle={selectedEventForRegistrations.Title}
        />
      )}
    </>
  );
};

export default AdminListView;

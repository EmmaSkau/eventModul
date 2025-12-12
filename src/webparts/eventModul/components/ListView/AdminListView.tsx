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
  DefaultButton,
  ActionButton,
} from "@fluentui/react";
import CreateEvent from "../CreateEvent";
import ConfirmDialog from "../Utility/ConfirmDialog";
import { IListViewProps } from "../Utility/IListViewProps";
import { IEventItem } from "../Utility/IEventItem";
import {
  getFutureEventsSorted,
  getPastEventsSorted,
  formatDate,
} from "../Utility/formatDate";
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
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [confirmDialogConfig, setConfirmDialogConfig] = useState<{
    title: string;
    message: string;
    onConfirm: () => Promise<void>;
  }>({ title: "", message: "", onConfirm: async () => {} });
  const [successMessage, setSuccessMessage] = useState<string | undefined>();
  const [errorMessage, setErrorMessage] = useState<string | undefined>();

  // Load registration counts
  const loadRegistrationCounts = useCallback(async (): Promise<void> => {
    try {
      const sp = getSP(props.context);

      // Get only registered users (not waitlist)
      const timestamp = Date.now();
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(
          `EventType eq 'Registered' and (Id ge 0 or Id eq ${timestamp})`
        )
        .select("EventId")
        .top(5000)();

      // Count registrations per event
      const counts: { [eventId: number]: number } = {};

      registrations.forEach((reg: { EventId?: number }) => {
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

      // Show past or future events based on toggle
      if (props.showPastEvents) {
        filtered = getPastEventsSorted(filtered);
      } else {
        filtered = getFutureEventsSorted(filtered);
      }

      return filtered;
    },
    [props.startDate, props.endDate, props.selectedLocation, props.showPastEvents]
  );

  // Load events
  const loadEvents = useCallback(async (): Promise<void> => {
    try {
      setIsLoading(true);
      setError(undefined);
      const sp = getSP(props.context);
      const currentUserEmail = props.context.pageContext.user.email;

      const timestamp = Date.now();
      const items: IEventItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.select(
          "Id",
          "Title",
          "Dato",
          "SlutDato",
          "Administrator/Title",
          "Placering",
          "Capacity",
          "M_x00e5_lgruppe/Title"
        )
        .expand("Administrator", "M_x00e5_lgruppe")
        .filter(
          `Administrator/EMail eq '${currentUserEmail}' and (Id ge 0 or Id eq ${timestamp})`
        )
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

  // Load data on mount and when filters change
  useEffect(() => {
    loadRegistrationCounts().catch(console.error);
    loadEvents().catch(console.error);
  }, [props.startDate, props.endDate, props.selectedLocation, props.showPastEvents]);

  // Event handlers
  const handleEditEvent = (item: IEventItem): void => {
    setEditPanelOpen(true);
    setSelectedEventForEdit(item);
  };

  const handleCloseEditPanel = (): void => {
    setEditPanelOpen(false);
    setSelectedEventForEdit(undefined);
  };

  const handleDeleteEvent = (item: IEventItem): void => {
    setConfirmDialogConfig({
      title: `Slet "${item.Title}"`,
      message:
        "Er du sikker på, at du vil slette dette event? Alle tilmeldinger og felter vil også blive slettet.",
      onConfirm: async () => {
        setShowConfirmDialog(false);
        setSuccessMessage(undefined);
        setErrorMessage(undefined);

        try {
          const sp = getSP(props.context);

          // Delete all registrations
          const registrations = await sp.web.lists
            .getByTitle("EventRegistrations")
            .items.filter(`EventId eq ${item.Id}`)
            .select("Id")();

          // Delete all event fields
          const eventFields = await sp.web.lists
            .getByTitle("EventFields")
            .items.filter(`EventId eq ${item.Id}`)
            .select("Id")();

          // Delete all registrations and fields in parallel using Promise.all
          await Promise.all([
            ...registrations.map((registration) =>
              sp.web.lists
                .getByTitle("EventRegistrations")
                .items.getById(registration.Id)
                .delete()
            ),
            ...eventFields.map((field) =>
              sp.web.lists
                .getByTitle("EventFields")
                .items.getById(field.Id)
                .delete()
            ),
          ]);

          // Delete the event itself
          await sp.web.lists
            .getByTitle("EventDB")
            .items.getById(item.Id)
            .delete();

          setSuccessMessage("Event slettet!");

          await loadEvents();
          await loadRegistrationCounts();
        } catch (error) {
          console.error("Error deleting event:", error);
          setErrorMessage("Fejl ved sletning af event. Prøv igen.");
        }
      },
    });
    setShowConfirmDialog(true);
  };

  const openManageRegistrations = (item: IEventItem): void => {
    setManageRegistrationsOpen(true);
    setSelectedEventForRegistrations(item);
  };

  const closeManageRegistrations = (): void => {
    setManageRegistrationsOpen(false);
    setSelectedEventForRegistrations(undefined);
    loadRegistrationCounts().catch(console.error);
    loadEvents().catch(console.error);
  };

  const getColumns = (): IColumn[] => {
    return [
      {
        key: "Title",
        name: "Titel",
        fieldName: "Title",
        minWidth: 200,
        maxWidth: 250,
        isResizable: true,
      },
      {
        key: "StartDato",
        name: "Start Dato",
        fieldName: "Dato",
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.Dato ? formatDate(item.Dato) : "-";
        },
      },
      {
        key: "SlutDato",
        name: "Slut Dato",
        fieldName: "SlutDato",
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return item.SlutDato ? formatDate(item.SlutDato) : "-";
        },
      },
      {
        key: "Placering",
        name: "Placering",
        fieldName: "Placering",
        minWidth: 100,
        maxWidth: 120,
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
        minWidth: 120,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IEventItem) => {
          if (item.M_x00e5_lgruppe && Array.isArray(item.M_x00e5_lgruppe)) {
            return item.M_x00e5_lgruppe.map((user: { Title: string }) => user.Title).join(", ");
          }
          return "-";
        },
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
              <DefaultButton
                text="Slet"
                onClick={() => handleDeleteEvent(item)}
                styles={{
                  root: { color: "#a4262c" },
                  rootHovered: { color: "#8c1c1e" },
                }}
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
      {successMessage && (
        <MessageBar
          messageBarType={MessageBarType.success}
          onDismiss={() => setSuccessMessage(undefined)}
        >
          {successMessage}
        </MessageBar>
      )}
      {errorMessage && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setErrorMessage(undefined)}
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

      <ConfirmDialog
        hidden={!showConfirmDialog}
        title={confirmDialogConfig.title}
        message={confirmDialogConfig.message}
        onConfirm={confirmDialogConfig.onConfirm}
        onCancel={() => setShowConfirmDialog(false)}
        confirmText="Slet"
        cancelText="Annuller"
      />
    </>
  );
};

export default AdminListView;

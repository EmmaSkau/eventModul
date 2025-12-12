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
} from "@fluentui/react";
import { IListViewProps } from "../Utility/IListViewProps";
import { IEventItem } from "../Utility/IEventItem";
import { getFutureEventsSorted, formatDate } from "../Utility/formatDate";
import ConfirmDialog from "../Utility/ConfirmDialog";

const RegisteredListView: React.FC<IListViewProps> = (props) => {
  // State declarations
  const [events, setEvents] = useState<IEventItem[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [confirmDialogConfig, setConfirmDialogConfig] = useState<{
    title: string;
    message: string;
    onConfirm: () => Promise<void>;
  }>({ title: "", message: "", onConfirm: async () => {} });
  const [successMessage, setSuccessMessage] = useState<string | undefined>();
  const [errorMessage, setErrorMessage] = useState<string | undefined>();

  // Filter events
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
      const currentUser = await sp.web.currentUser();

      // Build filter based on waitlisted toggle
      let registrationFilter = `BrugerId eq ${currentUser.Id}`;
      if (props.waitlisted) {
        registrationFilter += " and EventType eq 'Waitlist'";
      } else if (props.registered) {
        registrationFilter += " and EventType ne 'Waitlist'";
      }

      const timestamp = Date.now();
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(
          `${registrationFilter} and (Id ge 0 or Id eq ${timestamp})`
        )
        .select("EventId", "EventType")
        .top(5000)();

      const registeredEventIds = registrations
        .map((reg) => reg.EventId)
        .filter((id) => id);

      if (registeredEventIds.length === 0) {
        setEvents([]);
        setIsLoading(false);
        return;
      }

      const idFilters = registeredEventIds
        .map((id) => `Id eq ${id}`)
        .join(" or ");

      const items: IEventItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.filter(`(${idFilters}) and (Id ge 0 or Id eq ${timestamp})`)
        .select(
          "Id",
          "Title",
          "Dato",
          "SlutDato",
          "Administrator/Title",
          "Placering",
          "Capacity",
          "Online"
        )
        .expand("Administrator")
        .top(1000)();

      const filteredItems = filterEvents(items);

      setEvents(filteredItems);
      setIsLoading(false);
    } catch (error) {
      console.error("Error loading registered events:", error);
      setIsLoading(false);
      setError("Kunne ikke indlæse dine tilmeldte events");
    }
  }, [props.context, props.waitlisted, props.registered, filterEvents]);

  // Load data on mount
  useEffect(() => {
    loadEvents().catch(console.error);
  }, []);

  // Reload when filters change
  useEffect(() => {
    loadEvents().catch(console.error);
  }, [loadEvents]);

  // Handle unregister from event
  const handleDeleteEvent = (item: IEventItem): void => {
    setConfirmDialogConfig({
      title: `Afmeld "${item.Title}"`,
      message: "Er du sikker på, at du vil afmeldes dette event?",
      onConfirm: async () => {
        setShowConfirmDialog(false);
        setSuccessMessage(undefined);
        setErrorMessage(undefined);

        try {
          const sp = getSP(props.context);
          const currentUser = await sp.web.currentUser();
          const timestamp = Date.now();
          const registrations = await sp.web.lists
            .getByTitle("EventRegistrations")
            .items.filter(
              `BrugerId eq ${currentUser.Id} and EventId eq ${item.Id} and (Id ge 0 or Id eq ${timestamp})`
            )
            .select("Id")
            .top(5000)();

          if (registrations.length === 0) {
            setErrorMessage("Kunne ikke finde din tilmelding.");
            return;
          }

          await sp.web.lists
            .getByTitle("EventRegistrations")
            .items.getById(registrations[0].Id)
            .delete();

          setSuccessMessage("Du er nu afmeldt eventet!");

          await loadEvents();
        } catch (error) {
          console.error("Error unregistering from event:", error);
          setErrorMessage("Fejl ved afmelding af event. Prøv igen.");
        }
      },
    });
    setShowConfirmDialog(true);
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
        key: "OnlineLink",
        name: "Online Link",
        fieldName: "Online",
        minWidth: 120,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IEventItem) => {
          if (item.Online?.Url) {
            return (
              <a
                href={item.Online.Url}
                target="_blank"
                rel="noopener noreferrer"
              >
                {item.Online.Description || "Join Online"}
              </a>
            );
          }
          return "-";
        },
      },
      {
        key: "actions",
        name: "Afmeld",
        fieldName: "actions",
        minWidth: 200,
        maxWidth: 250,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return (
            <DefaultButton
              text="Afmeld event"
              onClick={() => handleDeleteEvent(item)}
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

      <ConfirmDialog
        hidden={!showConfirmDialog}
        title={confirmDialogConfig.title}
        message={confirmDialogConfig.message}
        onConfirm={confirmDialogConfig.onConfirm}
        onCancel={() => setShowConfirmDialog(false)}
        confirmText="Afmeld"
        cancelText="Annuller"
      />
    </>
  );
};

export default RegisteredListView;

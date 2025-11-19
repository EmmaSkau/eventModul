import * as React from "react";
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
import { IListViewState } from "../Utility/IListViewState";
import { getFutureEventsSorted, formatDate } from "../Utility/formatDate";

export default class RegisteredListView extends React.Component<
  IListViewProps,
  IListViewState
> {
  constructor(props: IListViewProps) {
    super(props);
    this.state = {
      events: [],
      isLoading: false,
    };
  }

  public componentDidMount(): void {
    this.loadEvents().catch(console.error);
  }

  public componentDidUpdate(prevProps: IListViewProps): void {
    // Reload when filters change
    if (
      prevProps.startDate !== this.props.startDate ||
      prevProps.endDate !== this.props.endDate ||
      prevProps.selectedLocation !== this.props.selectedLocation
    ) {
      this.loadEvents().catch(console.error);
    }
  }

  public loadEvents = async (): Promise<void> => {
    try {
      this.setState({ isLoading: true, error: undefined });
      const sp = getSP(this.props.context);
      const currentUser = await sp.web.currentUser();

      // Step 1: Get all EventIds where the current user is registered
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`Title eq '${currentUser.Title}'`)
        .select("EventId")();

      // Extract the event IDs
      const registeredEventIds = registrations
        .map((reg) => reg.EventId)
        .filter((id) => id);

      if (registeredEventIds.length === 0) {
        // User has no registrations
        this.setState({
          events: [],
          isLoading: false,
        });
        return;
      }

      // Step 2: Get the events that match the registered event IDs
      // Build filter for multiple IDs: (Id eq 1) or (Id eq 2) or (Id eq 3)
      const idFilters = registeredEventIds
        .map((id) => `Id eq ${id}`)
        .join(" or ");

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
        .filter(idFilters)
        .top(1000)();

      // Apply additional filters from props (date range, location, etc.)
      const filteredItems = this.filterEvents(items);

      this.setState({
        events: filteredItems,
        isLoading: false,
      });
    } catch (error) {
      console.error("Error loading registered events:", error);
      this.setState({
        isLoading: false,
        error: "Kunne ikke indlæse dine tilmeldte events",
      });
    }
  };

  private filterEvents = (items: IEventItem[]): IEventItem[] => {
    let filtered = [...items];

    // Filter by start date
    if (this.props.startDate) {
      filtered = filtered.filter((item) => {
        if (!item.Dato) return false;
        const eventDate = new Date(item.Dato);
        return eventDate >= this.props.startDate!;
      });
    }

    // Filter by end date
    if (this.props.endDate) {
      filtered = filtered.filter((item) => {
        if (!item.Dato) return false;
        const eventDate = new Date(item.Dato);
        return eventDate <= this.props.endDate!;
      });
    }

    // Filter by location
    if (this.props.selectedLocation && this.props.selectedLocation !== "all") {
      filtered = filtered.filter((item) => {
        if (!item.Placering) return false;

        // Parse JSON to get DisplayName
        try {
          const parsed = JSON.parse(item.Placering);
          return parsed.DisplayName === this.props.selectedLocation;
        } catch {
          // If not JSON, compare directly
          return item.Placering === this.props.selectedLocation;
        }
      });
    }

    filtered = getFutureEventsSorted(filtered);

    return filtered;
  };

  private handleDeleteEvent = async (item: IEventItem): Promise<void> => {
    if (
      !confirm(`Er du sikker på, at du vil afmeldes "${item.Title}" event?`)
    ) {
      return;
    }

    try {
      const sp = getSP(this.props.context);
      const currentUser = await sp.web.currentUser();
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(
          `Title eq '${currentUser.Title}' and EventId eq ${item.Id}`
        )
        .select("Id")();

      if (registrations.length === 0) {
        alert("Kunne ikke finde din tilmelding.");
        return;
      }

      await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.getById(registrations[0].Id)
        .delete();

      alert("Du er nu afmeldt eventet!");

      await this.loadEvents();
    } catch (error) {
      console.error("Error unregistering from event:", error);
      alert("Fejl ved afmelding af event. Prøv igen.");
    }
  };

  private getColumns = (): IColumn[] => {
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
        key: "online",
        name: "Online",
        fieldName: "online",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
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
              onClick={() => this.handleDeleteEvent(item)}
            />
          );
        },
      },
    ];
  };

  public render(): React.ReactElement<IListViewProps> {
    const { events, isLoading, error } = this.state;

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
          columns={this.getColumns()}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      </>
    );
  }
}

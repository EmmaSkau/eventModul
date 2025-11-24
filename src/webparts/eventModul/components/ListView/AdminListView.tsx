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
  PrimaryButton,
  DefaultButton,
  ActionButton,
} from "@fluentui/react";
import CreateEvent from "../CreateEvent";
import { IListViewProps } from "../Utility/IListViewProps";
import { IEventItem } from "../Utility/IEventItem";
import { IListViewState as BaseListViewState } from "../Utility/IListViewState";
import { getFutureEventsSorted, formatDate } from "../Utility/formatDate";
import ManageRegistrations from "../ManageRegistrations";

interface IListViewState extends BaseListViewState {
  editPanelOpen: boolean;
  selectedEventForEdit?: IEventItem;
  registrationCounts: { [eventId: number]: number };
  manageRegistrationsOpen: boolean;
  selectedEventForRegistrations?: IEventItem;
}

export default class AdminListView extends React.Component<
  IListViewProps,
  IListViewState
> {
  constructor(props: IListViewProps) {
    super(props);
    this.state = {
      events: [],
      isLoading: false,
      editPanelOpen: false,
      registrationCounts: {},
      manageRegistrationsOpen: false,
      selectedEventForRegistrations: undefined,
    };
  }

  public componentDidMount(): void {
    this.loadRegistrationCounts().catch(console.error);
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

  private loadRegistrationCounts = async (): Promise<void> => {
    try {
      const sp = getSP(this.props.context);

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

      this.setState({ registrationCounts: counts });
    } catch (error) {
      console.error("Error loading registration counts:", error);
    }
  };

  public loadEvents = async (): Promise<void> => {
    try {
      this.setState({ isLoading: true, error: undefined });
      const sp = getSP(this.props.context);
      const currentUserEmail = this.props.context.pageContext.user.email;

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

      const filteredItems = this.filterEvents(items);

      this.setState({
        events: filteredItems,
        isLoading: false,
      });
    } catch (error) {
      console.error("Error loading events:", error);
      this.setState({
        isLoading: false,
        error: "Kunne ikke indlæse events fra SharePoint",
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

  private handleEditEvent = (item: IEventItem): void => {
    this.setState({
      editPanelOpen: true,
      selectedEventForEdit: item,
    });
  };

  private handleCloseEditPanel = (): void => {
    this.setState({
      editPanelOpen: false,
      selectedEventForEdit: undefined,
    });
    this.loadRegistrationCounts().catch(console.error);
  };

  private handleDeleteEvent = async (item: IEventItem): Promise<void> => {
    // Confirm before deleting
    if (!confirm(`Er du sikker på, at du vil slette "${item.Title}"?`)) {
      return;
    }

    try {
      const sp = getSP(this.props.context);

      await sp.web.lists.getByTitle("EventDB").items.getById(item.Id).delete();

      alert("Event slettet!");

      // Reload the list
      await this.loadEvents();
      await this.loadRegistrationCounts();
    } catch (error) {
      console.error("Error deleting event:", error);
      alert("Fejl ved sletning af event. Prøv igen.");
    }
  };

  private openManageRegistrations = (item: IEventItem): void => {
    this.setState({
      manageRegistrationsOpen: true,
      selectedEventForRegistrations: item,
    });
  };

  private closeManageRegistrations = (): void => {
    this.setState({
      manageRegistrationsOpen: false,
      selectedEventForRegistrations: undefined,
    });
    // Refresh counts after managing registrations
    this.loadRegistrationCounts().catch(console.error);
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
        minWidth: 80,
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
        minWidth: 80,
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
        minWidth: 80,
        maxWidth: 150,
        isResizable: true,
      },
      {
        key: "tilmeldt",
        name: "Tilmeldte",
        fieldName: "tilmeldt",
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: IEventItem) => {
          const count = this.state.registrationCounts[item.Id] || 0;
          const capacity = item.Capacity || 0;
          const displayText =
            capacity > 0 ? `${count}/${capacity}` : count.toString();
          return (
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span>{displayText}</span>
              <ActionButton
                iconProps={{ iconName: "Edit" }}
                title="Administrer tilmeldte"
                onClick={() => this.openManageRegistrations(item)}
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
        maxWidth: 250,
        isResizable: true,
        onRender: (item: IEventItem) => {
          return (
            <div style={{ display: "flex", gap: "8px" }}>
              <PrimaryButton
                text="Ret"
                onClick={() => this.handleEditEvent(item)}
              />
              <DefaultButton
                text="Slet"
                onClick={() => this.handleDeleteEvent(item)}
              />
            </div>
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

        <CreateEvent
          isOpen={this.state.editPanelOpen}
          onClose={this.handleCloseEditPanel}
          context={this.props.context}
          onEventCreated={this.loadEvents}
          eventToEdit={this.state.selectedEventForEdit}
        />

        {this.state.manageRegistrationsOpen &&
          this.state.selectedEventForRegistrations && (
            <ManageRegistrations
              isOpen={this.state.manageRegistrationsOpen}
              onDismiss={this.closeManageRegistrations}
              context={this.props.context}
              eventId={this.state.selectedEventForRegistrations.Id}
              eventTitle={this.state.selectedEventForRegistrations.Title}
            />
          )}
      </>
    );
  }
}

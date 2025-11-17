import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "../../../pnpConfig";
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
} from "@fluentui/react";
import CreateEvent from "./CreateEvent";

export interface IListViewProps {
  context: WebPartContext;
  startDate?: Date;
  endDate?: Date;
  selectedLocation?: string;
  registered?: boolean;
  cancelledEvents?: boolean;
  waitlisted?: boolean;
}

interface IEventItem {
  Id: number;
  Title: string;
  Dato?: string; // StartDato internal name is "Dato"
  SlutDato?: string;
  Administrator?: {
    Title: string;
  };
  Placering?: string;
  targetGroup?: string;
  Beskrivelse?: string;
  TilfoejEkstraInfo?: string;
  Capacity?: number;
}

interface IListViewState {
  events: IEventItem[];
  isLoading: boolean;
  error?: string;
  selectedEventId?: number;
  selectedEventTitle?: string;
  editPanelOpen: boolean;
  selectedEventForEdit?: IEventItem;
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
    } catch (error) {
      console.error("Error deleting event:", error);
      alert("Fejl ved sletning af event. Prøv igen.");
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
          return item.Dato ? new Date(item.Dato).toLocaleDateString("DK") : "-";
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
          return item.SlutDato
            ? new Date(item.SlutDato).toLocaleDateString("DK")
            : "-";
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
        key: "actions",
        name: "Handlinger",
        fieldName: "actions",
        minWidth: 200,
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
    const { events, isLoading, error, } =
      this.state;

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
      </>
    );
  }
}

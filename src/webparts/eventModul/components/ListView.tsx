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
} from "@fluentui/react";
import RegisterForEvents from "./RegisterForEvents";

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
  Aldersgruppe?: string;
  Beskrivelse?: string;
  TilfoejEkstraInfo?: string;
  Capacity?: number;
}

interface IListViewState {
  events: IEventItem[];
  isLoading: boolean;
  error?: string;
  registerPanelOpen: boolean;
  selectedEventId?: number;
  selectedEventTitle?: string;
}

export default class ListView extends React.Component<IListViewProps, IListViewState> {
  constructor(props: IListViewProps) {
    super(props);
    this.state = {
      events: [],
      isLoading: false,
      registerPanelOpen: false,
    };
  }

  public componentDidMount(): void {
    this.loadEvents().catch(console.error);
  }

  public componentDidUpdate(prevProps: IListViewProps): void {
    // Reload when filters change
    if (prevProps.startDate !== this.props.startDate ||
        prevProps.endDate !== this.props.endDate ||
        prevProps.selectedLocation !== this.props.selectedLocation) {
      this.loadEvents().catch(console.error);
    }
  }

  private loadEvents = async (): Promise<void> => {
    try {
      this.setState({ isLoading: true, error: undefined });
      const sp = getSP(this.props.context);

      // Get items with just basic fields to test
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
        .expand("Administrator")();

      // Filter items based on props
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
      filtered = filtered.filter(item => {
        if (!item.Dato) return false;
        const eventDate = new Date(item.Dato);
        return eventDate >= this.props.startDate!;
      });
    }

    // Filter by end date
    if (this.props.endDate) {
      filtered = filtered.filter(item => {
        if (!item.Dato) return false;
        const eventDate = new Date(item.Dato);
        return eventDate <= this.props.endDate!;
      });
    }

    // Filter by location
    if (this.props.selectedLocation && this.props.selectedLocation !== 'all') {
      filtered = filtered.filter(item => 
        item.Placering === this.props.selectedLocation
      );
    }

    return filtered;
  };

  private openRegisterPanel = (eventId: number, eventTitle: string): void => {
    this.setState({
      registerPanelOpen: true,
      selectedEventId: eventId,
      selectedEventTitle: eventTitle,
    });
  };

  private closeRegisterPanel = (): void => {
    this.setState({
      registerPanelOpen: false,
      selectedEventId: undefined,
      selectedEventTitle: undefined,
    });
    // Reload events to reflect any changes
    this.loadEvents().catch(console.error);
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
          return item.SlutDato ? new Date(item.SlutDato).toLocaleDateString("DK") : "-";
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
      },
      {
        key: "Aldersgruppe",
        name: "Aldersgruppe",
        fieldName: "Aldersgruppe",
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
          return (
            <PrimaryButton
              text="Tilmeld"
              onClick={() => this.openRegisterPanel(item.Id, item.Title)}
            />
          );
        }
      }
    ];
  };

  public render(): React.ReactElement<IListViewProps> {
    const { events, isLoading, error, registerPanelOpen, selectedEventId, selectedEventTitle } = this.state;

    if (isLoading) {
      return <Spinner size={SpinnerSize.large} label="Indlæser events..." />;
    }

    if (error) {
      return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
    }

    if (events.length === 0) {
      return <MessageBar messageBarType={MessageBarType.info}>Ingen events fundet</MessageBar>;
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

        {registerPanelOpen && selectedEventId && selectedEventTitle && (
          <RegisterForEvents
            context={this.props.context}
            eventId={selectedEventId}
            eventTitle={selectedEventTitle}
            isOpen={registerPanelOpen}
            onDismiss={this.closeRegisterPanel}
          />
        )}
      </>
    );
  }
}
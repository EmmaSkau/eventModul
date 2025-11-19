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
} from "@fluentui/react";
import RegisterForEvents from "../RegisterForEvents";
import { IListViewProps } from "../Utility/IListViewProps";
import { IEventItem } from "../Utility/IEventItem";
import { IListViewState as BaseListViewState } from "../Utility/IListViewState";
import { getFutureEventsSorted, formatDate } from "../Utility/formatDate";

interface IListViewState extends BaseListViewState {
  registerPanelOpen: boolean;
  registeredEventIds: number[];
}

export default class ListView extends React.Component<
  IListViewProps,
  IListViewState
> {
  constructor(props: IListViewProps) {
    super(props);
    this.state = {
      events: [],
      isLoading: false,
      registerPanelOpen: false,
      registeredEventIds: [],
    };
  }

  public componentDidMount(): void {
    this.loadEvents().catch(console.error);
    this.loadUserRegistrations().catch(console.error); // Add this
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

      // Get ALL items by explicitly requesting a large number
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
        .top(1000)(); // Request up to 1000 items

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

  private closeRegisterPanel = (): void => {
    this.setState({
      registerPanelOpen: false,
      selectedEventId: undefined,
      selectedEventTitle: undefined,
    });
    this.loadEvents().catch(console.error);
  };

  // HANDLE REGISTRATIONS
  private checkIfEventHasCustomFields = async (
    eventId: number
  ): Promise<boolean> => {
    try {
      const sp = getSP(this.props.context);
      const fields = await sp.web.lists
        .getByTitle("EventFields")
        .items.filter(`EventId eq ${eventId}`)
        .select("Id")();

      return fields.length > 0;
    } catch (error) {
      console.error("Error checking custom fields:", error);
      return false;
    }
  };

  private registerUserToEvent = async (eventId: number): Promise<void> => {
    try {
      const sp = getSP(this.props.context);
      const currentUser = await sp.web.currentUser();

      // Generate a unique registration key
      const registrationKey = `${eventId}_${
        this.props.context.pageContext.user.loginName
      }_${new Date().getTime()}`;

      await sp.web.lists.getByTitle("EventRegistrations").items.add({
        Title: currentUser.Title,
        EventId: eventId,
        BrugerId: currentUser.Id,
        RegistrationKey: registrationKey,
        Submitted: new Date().toISOString(),
      });

      alert("Du er nu tilmeldt eventet!");

      await this.loadUserRegistrations();
      await this.loadEvents();
    } catch (error) {
      console.error("Error registering for event:", error);
      alert("Fejl ved tilmelding. Prøv igen.");
    }
  };

  private handleRegister = async (
    eventId: number,
    eventTitle: string
  ): Promise<void> => {
    const hasCustomFields = await this.checkIfEventHasCustomFields(eventId);

    if (hasCustomFields) {
      this.setState({
        registerPanelOpen: true,
        selectedEventId: eventId,
        selectedEventTitle: eventTitle,
      });
    } else {
      if (confirm(`Vil du tilmelde dig til "${eventTitle}"?`)) {
        await this.registerUserToEvent(eventId);
      }
    }
  };

  private loadUserRegistrations = async (): Promise<void> => {
    try {
      const sp = getSP(this.props.context);
      const currentUser = await sp.web.currentUser();

      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`Title eq '${currentUser.Title}'`)
        .select("EventId")();

      const registeredEventIds = registrations
        .map((reg) => reg.EventId)
        .filter((id) => id);

      this.setState({ registeredEventIds });
    } catch (error) {
      console.error("Error loading user registrations:", error);
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
          const isRegistered =
            this.state.registeredEventIds.indexOf(item.Id) !== -1;

          return (
            <PrimaryButton
              text={isRegistered ? "Tilmeldt" : "Tilmeld"}
              onClick={() => this.handleRegister(item.Id, item.Title)}
              disabled={isRegistered}
            />
          );
        },
      },
    ];
  };

  public render(): React.ReactElement<IListViewProps> {
    const {
      events,
      isLoading,
      error,
      registerPanelOpen,
      selectedEventId,
      selectedEventTitle,
    } = this.state;

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

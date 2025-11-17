import * as React from "react";
import { getSP } from "../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import {
  DatePicker,
  DayOfWeek,
  PrimaryButton,
  TextField,
  Dropdown,
  IDropdownOption,
  Panel,
  PanelType,
  DefaultButton,
  Stack,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IEventItem {
  Id: number;
  Title: string;
  Dato?: string;
  SlutDato?: string;
  Administrator?: {
    Title: string;
  };
  Placering?: string;
  Capacity?: number;
}

interface ICreateEventProps {
  onClose?: () => void;
  context: WebPartContext;
  onEventCreated?: () => void;
  eventToEdit?: IEventItem;
  isOpen: boolean;
}

interface ICreateEventState {
  // Form fields
  title?: string;
  startDate?: Date;
  endDate?: Date;
  selectedLocation?: string;
  maxParticipants?: number;

  // Loading states
  isLoadingEmployees: boolean;
  locationOptions: IDropdownOption[];
  isLoadingLocations: boolean;
  isSaving: boolean;
}

export default class CreateEvent extends React.Component<
  ICreateEventProps,
  ICreateEventState
> {
  constructor(props: ICreateEventProps) {
    super(props);
    const eventToEdit = props.eventToEdit;
    const isEditing = !!eventToEdit;

    this.state = {
      // Form fields
      title: isEditing ? eventToEdit!.Title : "",
      startDate:
        isEditing && eventToEdit!.Dato
          ? new Date(eventToEdit!.Dato)
          : undefined,
      endDate:
        isEditing && eventToEdit!.SlutDato
          ? new Date(eventToEdit!.SlutDato)
          : undefined,
      selectedLocation: isEditing ? eventToEdit!.Placering : undefined,
      maxParticipants:
        isEditing && eventToEdit!.Capacity ? eventToEdit!.Capacity : undefined,

      // Loading states
      isLoadingEmployees: false,
      locationOptions: [],
      isLoadingLocations: false,
      isSaving: false,
    };
  }

  public componentDidMount(): void {
    this.loadLocationsFromSharePoint().catch((error) => {
      console.error("Error loading locations:", error);
    });
  }

  public componentDidUpdate(prevProps: ICreateEventProps): void {
    if (
      this.props.eventToEdit &&
      this.props.eventToEdit !== prevProps.eventToEdit
    ) {
      const eventToEdit = this.props.eventToEdit;
      this.setState({
        title: eventToEdit.Title,
        startDate: eventToEdit.Dato ? new Date(eventToEdit.Dato) : undefined,
        endDate: eventToEdit.SlutDato
          ? new Date(eventToEdit.SlutDato)
          : undefined,
        selectedLocation: eventToEdit.Placering,
        maxParticipants: eventToEdit.Capacity,
      });
    }
    // When switching from edit to create mode, clear the form
    else if (!this.props.eventToEdit && prevProps.eventToEdit) {
      this.setState({
        title: "",
        startDate: undefined,
        endDate: undefined,
        selectedLocation: undefined,
        maxParticipants: undefined,
      });
    }
  }

  // LOCATIONS START
  private loadLocationsFromSharePoint = async (): Promise<void> => {
    try {
      this.setState({ isLoadingLocations: true });
      const sp = getSP(this.props.context);

      const items: { Placering: string }[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.select("Placering")();

      const locations = items
        .map((item) => item.Placering)
        .filter((location) => location);

      const uniqueLocations: string[] = [];
      const seen: { [key: string]: boolean } = {};
      for (const location of locations) {
        if (!seen[location]) {
          seen[location] = true;
          uniqueLocations.push(location);
        }
      }

      const locationOptions: IDropdownOption[] = uniqueLocations.map(
        (location) => ({
          key: location,
          text: location,
        })
      );

      this.setState({
        locationOptions,
        isLoadingLocations: false,
      });
    } catch (error) {
      console.error("Error loading locations from SharePoint:", error);
      this.setState({
        isLoadingLocations: false,
        locationOptions: [],
      });
    }
  };

  private onLocationChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ selectedLocation: option.key as string });
    }
  };

  // SAVE EVENT START
  private onTitleChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ title: newValue });
  };

  private onStartDateChange = (date: Date | null | undefined): void => {
    this.setState({ startDate: date || undefined });
  };

  private onEndDateChange = (date: Date | null | undefined): void => {
    this.setState({ endDate: date || undefined });
  };

  private saveEvent = async (): Promise<void> => {
    try {
      this.setState({ isSaving: true });

      const { title, startDate, endDate, selectedLocation, maxParticipants } =
        this.state;

      // Validation
      if (!title || !startDate || !endDate) {
        alert(
          "Please fill in all required fields (Title, Start Date, End Date)"
        );
        this.setState({ isSaving: false });
        return;
      }

      const sp = getSP(this.props.context);

      // Get current user's ID for the Administrator field
      const currentUser = await sp.web.currentUser();

      // Prepare the item data with CORRECT column names
      const itemData: {
        Title: string;
        Dato: string;
        SlutDato: string;
        AdministratorId: number;
        Placering: string;
        Capacity: number | null;
      } = {
        Title: title,
        Dato: startDate.toISOString(),
        SlutDato: endDate.toISOString(),
        AdministratorId: currentUser.Id, // Person field needs user ID (integer)
        Placering: selectedLocation || "",
        Capacity: maxParticipants
          ? parseInt(String(maxParticipants), 10)
          : null,
      };

      // Check if we're editing or creating
      if (this.props.eventToEdit) {
        // UPDATE existing item
        await sp.web.lists
          .getByTitle("EventDB")
          .items.getById(this.props.eventToEdit.Id)
          .update(itemData);
        alert("Event opdateret!");
      } else {
        // CREATE new item
        await sp.web.lists
          .getByTitle("EventDB")
          .items.add(itemData);
        alert("Event oprettet!");
      }

      // Notify parent that event was created/updated so ListView can refresh
      if (this.props.onEventCreated) {
        this.props.onEventCreated();
      }

      this.setState({ isSaving: false });

      // Close the form
      if (this.props.onClose) {
        this.props.onClose();
      }
    } catch (error) {
      console.error("Error saving event:", error);
      alert(
        this.props.eventToEdit
          ? "Fejl ved opdatering af event. Prøv igen."
          : "Fejl ved oprettelse af event. Prøv igen."
      );
      this.setState({ isSaving: false });
    }
  };

  public render(): React.ReactElement {
    const { isOpen, onClose, eventToEdit } = this.props;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onClose}
        type={PanelType.medium}
        headerText={eventToEdit ? "Ret event" : "Opret ny event"}
        closeButtonAriaLabel="Luk"
      >
        <Stack tokens={{ childrenGap: 15 }}>
            <DatePicker
              label="Fra"
              firstDayOfWeek={DayOfWeek.Monday}
              showWeekNumbers={true}
              placeholder="Vælg start dato"
              ariaLabel="Vælg start dato"
              value={this.state.startDate}
              onSelectDate={this.onStartDateChange}
            />
            <DatePicker
              label="Til"
              firstDayOfWeek={DayOfWeek.Monday}
              showWeekNumbers={true}
              placeholder="Vælg slut dato"
              ariaLabel="Vælg slut dato"
              value={this.state.endDate}
              onSelectDate={this.onEndDateChange}
            />

          <TextField
            label="Title"
            value={this.state.title}
            onChange={this.onTitleChange}
            required
          />

          <Dropdown
            label="Placering"
            placeholder={
              this.state.isLoadingLocations
                ? "Indlæser lokationer..."
                : "Vælg lokation..."
            }
            options={this.state.locationOptions}
            selectedKey={this.state.selectedLocation}
            onChange={this.onLocationChange}
            disabled={this.state.isLoadingLocations}
          />

          <TextField
            label="Kapacitet"
            type="number"
            value={this.state.maxParticipants?.toString() || ""}
            onChange={(e, newValue) => {
              const numValue = newValue ? parseInt(newValue, 10) : undefined;
              this.setState({ maxParticipants: numValue });
            }}
          />

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton
              text={eventToEdit ? "Gem ændringer" : "Gem event"}
              onClick={this.saveEvent}
              disabled={this.state.isSaving}
            />
            <DefaultButton
              text="Annuller"
              onClick={onClose}
              disabled={this.state.isSaving}
            />
          </Stack>
        </Stack>
      </Panel>
    );
  }
}

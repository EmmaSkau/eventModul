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
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from "./EventModul.module.scss";

interface ICreateEventProps {
  onClose?: () => void;
  context: WebPartContext;
  onEventCreated?: () => void;
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
    this.state = {
      // Form fields
      title: "",
      startDate: undefined,
      endDate: undefined,
      selectedLocation: undefined,
      maxParticipants: undefined,

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

      const {
        title,
        startDate,
        endDate,
        selectedLocation,
        maxParticipants,
      } = this.state;

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
      const itemData: any = {
        Title: title,
        Dato: startDate.toISOString(),
        SlutDato: endDate.toISOString(),
        AdministratorId: currentUser.Id, // Person field needs user ID (integer)
        Placering: selectedLocation || "",
        Capacity: maxParticipants ? parseInt(String(maxParticipants), 10) : null,
      };

      // Create the item
      const result = await sp.web.lists.getByTitle("EventDB").items.add(itemData);

      console.log("Event created successfully:", result);
      
      // Notify parent that event was created so ListView can refresh
      if (this.props.onEventCreated) {
        this.props.onEventCreated();
      }

      this.setState({ isSaving: false });
      
      alert("Event created successfully!");

      // Close the form
      if (this.props.onClose) {
        this.props.onClose();
      }
    } catch (error) {
      console.error("Error saving event:", error);
      alert("Error creating event. Please try again.");
      this.setState({ isSaving: false });
    }
  };

  public render(): React.ReactElement {
    return (
      <section>
        <h1>Opret ny event</h1>

        <div className={styles.filters}>
          <DatePicker
            label="Fra"
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            placeholder="Select start date"
            ariaLabel="Select start date"
            value={this.state.startDate}
            onSelectDate={this.onStartDateChange}
          />
          <DatePicker
            label="Til"
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            placeholder="Select end date"
            ariaLabel="Select end date"
            value={this.state.endDate}
            onSelectDate={this.onEndDateChange}
          />
        </div>

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
          onChange={(e, newValue) => {
            const numValue = newValue ? parseInt(newValue, 10) : undefined;
            this.setState({ maxParticipants: numValue });
          }}
        />

        <PrimaryButton
          text="Afslut og gem event"
          onClick={this.saveEvent}
          disabled={this.state.isSaving}
        />
      </section>
    );
  }
}

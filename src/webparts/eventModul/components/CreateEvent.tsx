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
  IDropdownOption,
  Panel,
  PanelType,
  DefaultButton,
  Stack,
  Label,
  IconButton,
  Text,
  Checkbox,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEventItem } from "../components/Utility/IEventItem";
import AddFieldDialog, { ICustomField } from "./SpecialFields";

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
  customFields: ICustomField[];
  isAddingField: boolean;
  showFieldDialog: boolean;
  isOnline: boolean;
  onlineLink?: string;

  // Loading states
  isLoadingEmployees: boolean;
  locationOptions: IDropdownOption[];
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
      customFields: [],
      isAddingField: false,
      showFieldDialog: false,
      isOnline: false,
      onlineLink: "",

      // Loading states
      isLoadingEmployees: false,
      locationOptions: [],
      isSaving: false,
    };
  }

  public componentDidMount(): void {
    this.loadLocationsFromSharePoint().catch((error) => {
      console.error("Error loading locations:", error);
    });

    // Load custom fields if editing an existing event
    if (this.props.eventToEdit) {
      this.loadCustomFields(this.props.eventToEdit.Id).catch((error) => {
        console.error("Error loading custom fields:", error);
      });
    }
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
        isOnline: eventToEdit.Placering === "Online",
        onlineLink: eventToEdit.Online?.Url || "",
      });

      // Load custom fields for the new event being edited
      this.loadCustomFields(eventToEdit.Id).catch((error) => {
        console.error("Error loading custom fields:", error);
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
        customFields: [],
        isOnline: false,
        onlineLink: "",
      });
    }
  }

  // LOCATIONS START
  private loadLocationsFromSharePoint = async (): Promise<void> => {
    try {
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
      });
    } catch (error) {
      console.error("Error loading locations from SharePoint:", error);
      this.setState({
        locationOptions: [],
      });
    }
  };

  // ONLINE CHECKBOX START
  private onOnlineCheckboxChange = (
    event?: React.FormEvent<HTMLElement | HTMLInputElement>,
    checked?: boolean
  ): void => {
    this.setState({
      isOnline: !!checked,
      onlineLink: checked ? this.state.onlineLink : "",
    });
  };

  private onOnlineLinkChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ onlineLink: newValue });
  };

  // ADD CUSTOM FIELDS START
  private openAddFieldDialog = (): void => {
    this.setState({ showFieldDialog: true });
  };

  private addCustomField = (field: ICustomField): void => {
    this.setState((prevState) => ({
      customFields: [...prevState.customFields, field],
      showFieldDialog: false,
    }));
  };

  private cancelAddField = (): void => {
    this.setState({ showFieldDialog: false });
  };

  private removeCustomField = (fieldId: string): void => {
    this.setState((prevState) => ({
      customFields: prevState.customFields.filter((f) => f.id !== fieldId),
    }));
  };

  private loadCustomFields = async (eventId: number): Promise<void> => {
    try {
      const sp = getSP(this.props.context);
      const fields = await sp.web.lists
        .getByTitle("EventFields")
        .items.filter(`EventId eq ${eventId}`)
        .select("Id", "Title", "FeltType", "Valgmuligheder")();

      const customFields: ICustomField[] = fields.map((item) => ({
        id: item.Id.toString(),
        fieldName: item.Title,
        fieldType: item.FeltType,
        options: item.Valgmuligheder
          ? JSON.parse(item.Valgmuligheder)
          : undefined,
      }));

      this.setState({ customFields });
    } catch (error) {
      console.error("Error loading custom fields:", error);
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

  private onLocationChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ selectedLocation: newValue });
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
        Online?: {
          Description: string;
          Url: string;
        } | null;
      } = {
        Title: title,
        Dato: startDate.toISOString(),
        SlutDato: endDate.toISOString(),
        AdministratorId: currentUser.Id,
        Placering: this.state.isOnline ? "Online" : selectedLocation || "", 
        Capacity: maxParticipants
          ? parseInt(String(maxParticipants), 10)
          : null,
        Online:
          this.state.isOnline && this.state.onlineLink 
            ? {
                Description: "Online Link",
                Url: this.state.onlineLink,
              }
            : null,
      };

      // Check if we're editing or creating
      let eventId: number;
      if (this.props.eventToEdit) {
        await sp.web.lists
          .getByTitle("EventDB")
          .items.getById(this.props.eventToEdit.Id)
          .update(itemData);
        eventId = this.props.eventToEdit.Id;

        // Delete existing custom fields for this event
        const existingFields = await sp.web.lists
          .getByTitle("EventFields")
          .items.filter(`EventId eq ${eventId}`)
          .select("Id")();

        for (const existingField of existingFields) {
          await sp.web.lists
            .getByTitle("EventFields")
            .items.getById(existingField.Id)
            .delete();
        }

        alert("Event opdateret!");
      } else {
        // CREATE new item
        const addResult = await sp.web.lists
          .getByTitle("EventDB")
          .items.add(itemData);
        eventId = addResult.data?.Id || addResult.Id;
        alert("Event oprettet!");
      }

      // Save custom fields to EventFields list (only if there are any)
      if (this.state.customFields.length > 0) {
        for (const field of this.state.customFields) {
          await sp.web.lists.getByTitle("EventFields").items.add({
            Title: field.fieldName,
            EventId: eventId,
            FeltType: field.fieldType,
            Valgmuligheder: field.options
              ? JSON.stringify(field.options)
              : null,
          });
        }
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

          <TextField
            label="Placering"
            value={this.state.selectedLocation}
            onChange={this.onLocationChange}
            disabled={this.state.isOnline}
          />

          <Checkbox
            label="Online?"
            checked={this.state.isOnline}
            onChange={this.onOnlineCheckboxChange}
          />

          {this.state.isOnline && (
            <TextField
              label="Online Link (Teams/møde link)"
              placeholder="https://teams.microsoft.com/..."
              value={this.state.onlineLink}
              onChange={this.onOnlineLinkChange}
            />
          )}

          <TextField
            label="Kapacitet"
            type="number"
            value={this.state.maxParticipants?.toString() || ""}
            onChange={(e, newValue) => {
              const numValue = newValue ? parseInt(newValue, 10) : undefined;
              this.setState({ maxParticipants: numValue });
            }}
          />

          {this.state.showFieldDialog && (
            <AddFieldDialog
              onAddField={this.addCustomField}
              onCancel={this.cancelAddField}
            />
          )}

          {this.state.customFields.length > 0 && (
            <Stack tokens={{ childrenGap: 10 }}>
              <Label>Brugerdefinerede felter:</Label>
              {this.state.customFields.map((field) => (
                <Stack
                  key={field.id}
                  horizontal
                  horizontalAlign="space-between"
                >
                  <Text>
                    {field.fieldName} ({field.fieldType})
                  </Text>
                  <IconButton
                    iconProps={{ iconName: "Delete" }}
                    onClick={() => this.removeCustomField(field.id)}
                  />
                </Stack>
              ))}
            </Stack>
          )}

          <DefaultButton
            text="Tilføj flere felter"
            onClick={this.openAddFieldDialog}
            disabled={this.state.isSaving}
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

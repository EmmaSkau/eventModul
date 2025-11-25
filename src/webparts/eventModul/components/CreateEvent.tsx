import * as React from "react";
import { useState, useEffect, useCallback } from "react";
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

const CreateEvent: React.FC<ICreateEventProps> = (props) => {
  const { context, eventToEdit, onEventCreated, onClose, isOpen } = props;

  // Form fields state
  const [title, setTitle] = useState<string>("");
  const [startDate, setStartDate] = useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = useState<Date | undefined>(undefined);
  const [selectedLocation, setSelectedLocation] = useState<string | undefined>(
    undefined
  );
  const [maxParticipants, setMaxParticipants] = useState<number | undefined>(
    undefined
  );
  const [customFields, setCustomFields] = useState<ICustomField[]>([]);
  const [showFieldDialog, setShowFieldDialog] = useState(false);
  const [isOnline, setIsOnline] = useState(false);
  const [onlineLink, setOnlineLink] = useState<string>("");

  // Loading states
  const [isSaving, setIsSaving] = useState(false);

  const loadCustomFields = useCallback(
    async (eventId: number): Promise<void> => {
      try {
        const sp = getSP(context);
        const fields = await sp.web.lists
          .getByTitle("EventFields")
          .items.filter(`EventId eq ${eventId}`)
          .select(
            "Id",
            "Title",
            "FeltType",
            "Valgmuligheder",
            "P_x00e5_kr_x00e6_vet"
          )();

        setCustomFields(fields);
      } catch (error) {
        console.error("Error loading custom fields:", error);
      }
    },
    [context]
  );

  // Load custom fields when editing an event
  useEffect(() => {
    if (eventToEdit) {
      loadCustomFields(eventToEdit.Id).catch((error) => {
        console.error("Error loading custom fields:", error);
      });
    }
  }, [eventToEdit?.Id, loadCustomFields]);

  // Update form when eventToEdit changes
  useEffect(() => {
    if (eventToEdit) {
      setTitle(eventToEdit.Title);
      setStartDate(eventToEdit.Dato ? new Date(eventToEdit.Dato) : undefined);
      setEndDate(
        eventToEdit.SlutDato ? new Date(eventToEdit.SlutDato) : undefined
      );
      setSelectedLocation(eventToEdit.Placering);
      setMaxParticipants(eventToEdit.Capacity);
      setIsOnline(eventToEdit.Placering === "Online");
      setOnlineLink(eventToEdit.Online?.Url || "");
    } else {
      // Clear form when switching to create mode
      setTitle("");
      setStartDate(undefined);
      setEndDate(undefined);
      setSelectedLocation(undefined);
      setMaxParticipants(undefined);
      setCustomFields([]);
      setIsOnline(false);
      setOnlineLink("");
    }
  }, [eventToEdit]);

  // ONLINE CHECKBOX START
  const onOnlineCheckboxChange = useCallback(
    (
      event?: React.FormEvent<HTMLElement | HTMLInputElement>,
      checked?: boolean
    ): void => {
      setIsOnline(!!checked);
      if (!checked) {
        setOnlineLink("");
      }
    },
    []
  );

  const onOnlineLinkChange = useCallback(
    (
      event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setOnlineLink(newValue || "");
    },
    []
  );

  // ADD CUSTOM FIELDS START
  const openAddFieldDialog = useCallback((): void => {
    setShowFieldDialog(true);
  }, []);

  const addCustomField = useCallback((field: ICustomField): void => {
    setCustomFields((prev) => [...prev, field]);
    setShowFieldDialog(false);
  }, []);

  const cancelAddField = useCallback((): void => {
    setShowFieldDialog(false);
  }, []);

  const removeCustomField = useCallback((fieldId: string): void => {
    setCustomFields((prev) => prev.filter((f) => f.id !== fieldId));
  }, []);

  // FORM FIELD HANDLERS
  const onTitleChange = useCallback(
    (
      event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setTitle(newValue || "");
    },
    []
  );

  const onStartDateChange = useCallback(
    (date: Date | null | undefined): void => {
      setStartDate(date || undefined);
    },
    []
  );

  const onEndDateChange = useCallback((date: Date | null | undefined): void => {
    setEndDate(date || undefined);
  }, []);

  const onLocationChange = useCallback(
    (
      event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setSelectedLocation(newValue);
    },
    []
  );

  // SAVE EVENT
  const saveEvent = useCallback(async (): Promise<void> => {
    try {
      setIsSaving(true);

      // Validation
      if (!title || !startDate || !endDate) {
        alert(
          "Please fill in all required fields (Title, Start Date, End Date)"
        );
        setIsSaving(false);
        return;
      }

      const sp = getSP(context);

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
        Placering: isOnline ? "Online" : selectedLocation || "",
        Capacity: maxParticipants
          ? parseInt(String(maxParticipants), 10)
          : null,
        Online:
          isOnline && onlineLink
            ? {
                Description: "Online Link",
                Url: onlineLink,
              }
            : null,
      };

      // Check if we're editing or creating
      let eventId: number;
      if (eventToEdit) {
        await sp.web.lists
          .getByTitle("EventDB")
          .items.getById(eventToEdit.Id)
          .update(itemData);
        eventId = eventToEdit.Id;

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
      if (customFields.length > 0) {
        for (const field of customFields) {
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
      if (onEventCreated) {
        onEventCreated();
      }

      setIsSaving(false);

      // Close the form
      if (onClose) {
        onClose();
      }
    } catch (error) {
      console.error("Error saving event:", error);
      alert(
        eventToEdit
          ? "Fejl ved opdatering af event. Prøv igen."
          : "Fejl ved oprettelse af event. Prøv igen."
      );
      setIsSaving(false);
    }
  }, [
    title,
    startDate,
    endDate,
    context,
    isOnline,
    selectedLocation,
    maxParticipants,
    onlineLink,
    eventToEdit,
    customFields,
    onEventCreated,
    onClose,
  ]);

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
          value={startDate}
          onSelectDate={onStartDateChange}
        />
        <DatePicker
          label="Til"
          firstDayOfWeek={DayOfWeek.Monday}
          showWeekNumbers={true}
          placeholder="Vælg slut dato"
          ariaLabel="Vælg slut dato"
          value={endDate}
          onSelectDate={onEndDateChange}
        />

        <TextField
          label="Title"
          value={title}
          onChange={onTitleChange}
          required
        />

        <TextField
          label="Placering"
          value={selectedLocation}
          onChange={onLocationChange}
          disabled={isOnline}
        />

        <Checkbox
          label="Online?"
          checked={isOnline}
          onChange={onOnlineCheckboxChange}
        />

        {isOnline && (
          <TextField
            label="Online Link (Teams/møde link)"
            placeholder="https://teams.microsoft.com/..."
            value={onlineLink}
            onChange={onOnlineLinkChange}
          />
        )}

        <TextField
          label="Kapacitet"
          type="number"
          value={maxParticipants?.toString() || ""}
          onChange={(e, newValue) => {
            const numValue = newValue ? parseInt(newValue, 10) : undefined;
            setMaxParticipants(numValue);
          }}
        />

        {showFieldDialog && (
          <AddFieldDialog
            onAddField={addCustomField}
            onCancel={cancelAddField}
          />
        )}

        {customFields.length > 0 && (
          <Stack tokens={{ childrenGap: 10 }}>
            <Label>Brugerdefinerede felter:</Label>
            {customFields.map((field) => (
              <Stack key={field.id} horizontal horizontalAlign="space-between">
                <Text>
                  {field.fieldName} ({field.fieldType})
                </Text>
                <IconButton
                  iconProps={{ iconName: "Delete" }}
                  onClick={() => removeCustomField(field.id)}
                />
              </Stack>
            ))}
          </Stack>
        )}

        <DefaultButton
          text="Tilføj flere felter"
          onClick={openAddFieldDialog}
          disabled={isSaving}
        />

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text={eventToEdit ? "Gem ændringer" : "Gem event"}
            onClick={saveEvent}
            disabled={isSaving}
          />
          <DefaultButton
            text="Annuller"
            onClick={onClose}
            disabled={isSaving}
          />
        </Stack>
      </Stack>
    </Panel>
  );
};

export default CreateEvent;

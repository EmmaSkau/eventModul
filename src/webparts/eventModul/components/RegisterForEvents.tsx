import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  Panel,
  PanelType,
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Checkbox,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
} from "@fluentui/react";

export interface IRegisterForEventsProps {
  context: WebPartContext;
  eventId: number;
  eventTitle: string;
  isOpen: boolean;
  onDismiss: () => void;
}

interface IEventField {
  Id: number;
  Title: string;
  EventId: number;
  FeltType: string; // "text", "multipleChoice"
  Valgmuligheder?: string; // JSON string with options array
  P_x00e5_kr_x00e6_vet?: boolean; // "Påkrævet" (Required)
}

const RegisterForEvents: React.FC<IRegisterForEventsProps> = (props) => {
  const { context, eventId, eventTitle, isOpen, onDismiss } = props;

  // State
  const [fields, setFields] = useState<IEventField[]>([]);
  const [fieldValues, setFieldValues] = useState<{ [key: number]: string | boolean }>({});
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [success, setSuccess] = useState<string | undefined>(undefined);
  const [capacity, setCapacity] = useState<number>(0);
  const [registrationCount, setRegistrationCount] = useState<number>(0);

  const loadEventFields = useCallback(async (): Promise<void> => {
    try {
      setIsLoading(true);
      setError(undefined);
      const sp = getSP(context);
      
      // Load event details for capacity
      const event = await sp.web.lists
        .getByTitle("EventDB")
        .items.getById(eventId)
        .select("Capacity")();
      
      setCapacity(event.Capacity || 0);
      
      // Load current registration count (only 'Registered' status)
      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`EventId eq ${eventId} and EventType eq 'Registered'`)
        .select("Id")();
      
      setRegistrationCount(registrations.length);
      
      // Load custom fields
      const allItems = await sp.web.lists
        .getByTitle("EventFields")
        .items.select(
          "Id",
          "Title",
          "EventId",
          "FeltType",
          "Valgmuligheder",
          "P_x00e5_kr_x00e6_vet"
        )();

      const eventFields: IEventField[] = allItems.filter(
        (item) => item.EventId === eventId
      );

      setFields(eventFields);
      setIsLoading(false);
    } catch (error) {
      console.error("Error loading event fields:", error);
      setIsLoading(false);
      setError("Kunne ikke indlæse tilmeldingsfelter");
    }
  }, [context, eventId]);

  // Load event fields when panel opens
  useEffect(() => {
    if (isOpen) {
      loadEventFields().catch(console.error);
    }
  }, [isOpen, loadEventFields]);

  const onFieldChange = useCallback((fieldId: number, value: string | boolean): void => {
    setFieldValues((prev) => ({
      ...prev,
      [fieldId]: value,
    }));
  }, []);

  const validateForm = useCallback((): boolean => {
    for (const field of fields) {
      if (field.P_x00e5_kr_x00e6_vet) {
        const value = fieldValues[field.Id];
        if (value === undefined || value === "" || value === false) {
          setError(`Feltet "${field.Title}" er p\u00e5kr\u00e6vet`);
          return false;
        }
      }
    }

    return true;
  }, [fields, fieldValues]);

  const handleSubmit = useCallback(async (): Promise<void> => {
    if (!validateForm()) {
      return;
    }

    try {
      setIsSaving(true);
      setError(undefined);
      const sp = getSP(context);

      // Determine if user should be on waitlist
      const available = capacity - registrationCount;
      const isWaitlist = capacity > 0 && available <= 0;
      const eventType = isWaitlist ? "Waitlist" : "Registered";

      // Generate a unique registration key
      const registrationKey = `${eventId}_${
        context.pageContext.user.loginName
      }_${new Date().getTime()}`;

      const currentUser = await sp.web.currentUser();

      for (const field of fields) {
        const value = fieldValues[field.Id];
        if (value !== undefined) {
          const itemData = {
            Title: eventTitle,
            EventId: eventId, 
            BrugerId: currentUser.Id, 
            FieldName: field.Title,
            FieldType: field.FeltType,
            FieldValue: String(value),
            RegistrationKey: registrationKey,
            Submitted: new Date().toISOString(),
            EventType: eventType,
          };

          await sp.web.lists.getByTitle("EventRegistrations").items.add(itemData);
        }
      }

      setIsSaving(false);
      const successMessage = isWaitlist 
        ? "Du er tilføjet til ventelisten!" 
        : "Du er nu tilmeldt eventet!";
      setSuccess(successMessage);
      setFieldValues({});

      // Close panel after 2 seconds
      setTimeout(() => {
        onDismiss();
      }, 2000);
    } catch (error) {
      console.error("Error submitting registration:", error);
      setIsSaving(false);
      setError("Kunne ikke gemme tilmeldingen. Prøv igen.");
    }
  }, [validateForm, context, eventId, eventTitle, fields, fieldValues, capacity, registrationCount, onDismiss]);

  const renderField = useCallback((field: IEventField): JSX.Element => {
    const value = fieldValues[field.Id] || "";

    switch (field.FeltType) {
      case "text":
      case "Text":
      case "Tekst":
        return (
          <TextField
            label={field.Title}
            required={field.P_x00e5_kr_x00e6_vet}
            value={value as string}
            onChange={(_, newValue) =>
              onFieldChange(field.Id, newValue || "")
            }
          />
        );

      case "multipleChoice":
      case "Dropdown":
      case "Valgmenu": {
        // Parse the JSON string to get the options array
        let options: IDropdownOption[] = [];
        if (field.Valgmuligheder) {
          try {
            // Try to parse as JSON first (new format)
            const optionsArray = JSON.parse(field.Valgmuligheder);
            options = optionsArray.map((opt: string) => ({
              key: opt,
              text: opt,
            }));
          } catch {
            // Fallback to comma-separated (old format)
            options = field.Valgmuligheder.split(",").map((opt) => ({
              key: opt.trim(),
              text: opt.trim(),
            }));
          }
        }
        return (
          <Dropdown
            label={field.Title}
            required={field.P_x00e5_kr_x00e6_vet}
            options={options}
            selectedKey={value as string}
            onChange={(_, option) =>
              onFieldChange(field.Id, (option?.key as string) || "")
            }
            placeholder="Vælg en mulighed"
          />
        );
      }

      case "Checkbox":
      case "Afkrydsningsfelt":
        return (
          <Checkbox
            label={field.Title}
            checked={value as boolean}
            onChange={(_, checked) => onFieldChange(field.Id, !!checked)}
          />
        );

      default:
        console.warn("Unknown field type:", field.FeltType, "for field:", field.Title);
        return (
          <TextField
            label={field.Title}
            required={field.P_x00e5_kr_x00e6_vet}
            value={value as string}
            onChange={(_, newValue) =>
              onFieldChange(field.Id, newValue || "")
            }
          />
        );
    }
  }, [fieldValues, onFieldChange]);

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={`Tilmeld til: ${eventTitle}`}
      closeButtonAriaLabel="Luk"
    >
      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Indlæser felter..." />
      ) : (
        <Stack tokens={{ childrenGap: 15 }}>
          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}
          {success && (
            <MessageBar messageBarType={MessageBarType.success}>
              {success}
            </MessageBar>
          )}

          {fields.length === 0 && !error && (
            <MessageBar messageBarType={MessageBarType.info}>
              Ingen ekstra felter krævet for dette event
            </MessageBar>
          )}

          {fields.map((field) => (
            <div key={field.Id}>{renderField(field)}</div>
          ))}

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton
              text={isSaving ? "Gemmer..." : "Tilmeld"}
              onClick={handleSubmit}
              disabled={isSaving}
            />
            <DefaultButton
              text="Annuller"
              onClick={onDismiss}
              disabled={isSaving}
            />
          </Stack>
        </Stack>
      )}
    </Panel>
  );
};

export default RegisterForEvents;

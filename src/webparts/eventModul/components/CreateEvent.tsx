import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { getSP } from "../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import {
  DatePicker,
  TimePicker,
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
  SearchBox,
  Spinner,
  SpinnerSize,
  IComboBox,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEventItem } from "../components/Utility/IEventItem";
import AddFieldDialog, { ICustomField } from "./SpecialFields";
import styles from "./EventModul.module.scss";

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
  const [startTime, setStartTime] = useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = useState<Date | undefined>(undefined);
  const [endTime, setEndTime] = useState<Date | undefined>(undefined);
  const [sameDate, setSameDate] = useState<boolean>(false);
  const [selectedUsers, setSelectedUsers] = useState<
    Array<{ Id: number; Title: string; Email: string }>
  >([]);
  const [userSearchText, setUserSearchText] = useState<string>("");
  const [userSearchResults, setUserSearchResults] = useState<
    Array<{ Id: number; Title: string; Email: string }>
  >([]);
  const [isSearchingUsers, setIsSearchingUsers] = useState(false);
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

  // Search for users
  const searchUsers = useCallback(
    async (searchText: string): Promise<void> => {
      if (!searchText || searchText.length < 2) {
        setUserSearchResults([]);
        return;
      }

      try {
        setIsSearchingUsers(true);
        const sp = getSP(context);

        const users = await sp.web.siteUsers
          .filter(`substringof('${searchText}', Title)`)
          .top(10)();

        const results = users.map((user) => ({
          Id: user.Id,
          Title: user.Title,
          Email: user.Email || "",
        }));

        setUserSearchResults(results);
        setIsSearchingUsers(false);
      } catch (error) {
        console.error("Error searching users:", error);
        setIsSearchingUsers(false);
      }
    },
    [context]
  );

  const handleUserSearchChange = useCallback(
    (newValue: string | undefined): void => {
      const value = newValue || "";
      setUserSearchText(value);
      searchUsers(value).catch(console.error);
    },
    [searchUsers]
  );

  const addUserToTargetGroup = useCallback(
    (user: { Id: number; Title: string; Email: string }): void => {
      setSelectedUsers((prev) => {
        // Check if user is already added
        if (prev.some((u) => u.Id === user.Id)) {
          return prev;
        }
        return [...prev, user];
      });
      setUserSearchText("");
      setUserSearchResults([]);
    },
    []
  );

  const removeUserFromTargetGroup = useCallback((userId: number): void => {
    setSelectedUsers((prev) => prev.filter((u) => u.Id !== userId));
  }, []);

  const loadCustomFields = useCallback(
    async (eventId: number): Promise<void> => {
      try {
        const sp = getSP(context);
        // Force fresh data by using unique filter each time
        const timestamp = Date.now();
        const data = await sp.web.lists
          .getByTitle("EventFields")
          .items
          .filter(`EventId eq ${eventId} and (Id ge 0 or Id eq ${timestamp})`)
          .select("Id", "Title", "FeltType", "Valgmuligheder", "P_x00e5_kr_x00e6_vet")
          .top(5000)();
        
        setCustomFields(data);
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
      const startDateTime = eventToEdit.Dato
        ? new Date(eventToEdit.Dato)
        : undefined;
      const endDateTime = eventToEdit.SlutDato
        ? new Date(eventToEdit.SlutDato)
        : undefined;

      setStartDate(startDateTime);
      setStartTime(startDateTime);
      setEndDate(endDateTime);
      setEndTime(endDateTime);

      setSelectedLocation(eventToEdit.Placering);
      setMaxParticipants(eventToEdit.Capacity);
      setIsOnline(eventToEdit.Placering === "Online");
      setOnlineLink(eventToEdit.Online?.Url || "");

      // Load target group users if available
      if (
        eventToEdit.M_x00e5_lgruppeId &&
        eventToEdit.M_x00e5_lgruppeId.length > 0
      ) {
        const loadTargetGroupUsers = async (): Promise<void> => {
          try {
            const sp = getSP(context);
            const users = await Promise.all(
              eventToEdit.M_x00e5_lgruppeId!.map(async (userId: number) => {
                const user = await sp.web.siteUsers.getById(userId)();
                return {
                  Id: user.Id,
                  Title: user.Title,
                  Email: user.Email || "",
                };
              })
            );
            setSelectedUsers(users);
          } catch (error) {
            console.error("Error loading target group users:", error);
          }
        };
        loadTargetGroupUsers().catch(console.error);
      } else {
        setSelectedUsers([]);
      }
    } else {
      // Clear form when switching to create mode
      setTitle("");
      setStartDate(undefined);
      setStartTime(undefined);
      setEndDate(undefined);
      setEndTime(undefined);
      setSameDate(false);
      setSelectedLocation(undefined);
      setMaxParticipants(undefined);
      setCustomFields([]);
      setIsOnline(false);
      setOnlineLink("");
      setSelectedUsers([]);
      setUserSearchText("");
      setUserSearchResults([]);
    }
  }, [eventToEdit, context]);

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
      // If same day is checked, update end date to match start date
      if (sameDate && date) {
        setEndDate(date);
      }
    },
    [sameDate]
  );

  const onStartTimeChange = useCallback(
    (_event: React.FormEvent<IComboBox>, time: Date): void => {
      setStartTime(time);
    },
    []
  );

  const onEndTimeChange = useCallback(
    (_event: React.FormEvent<IComboBox>, time: Date): void => {
      setEndTime(time);
    },
    []
  );

  const onEndDateChange = useCallback((date: Date | null | undefined): void => {
    setEndDate(date || undefined);
  }, []);

  const sameDateChange = useCallback(
    (
      event?: React.FormEvent<HTMLElement | HTMLInputElement>,
      checked?: boolean
    ): void => {
      setSameDate(!!checked);
      if (checked && startDate) {
        setEndDate(startDate);
      }
    },
    [startDate]
  );

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
    setIsSaving(true);
    
    try {
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

      const combineDateAndTime = (date: Date, time?: Date): Date => {
        const combined = new Date(date.getTime());
        if (time) {
          const hours = time.getHours();
          const minutes = time.getMinutes();
          combined.setHours(hours, minutes, 0, 0);
        }
        return combined;
      };

      const finalStartDate = combineDateAndTime(startDate, startTime);
      const finalEndDate = combineDateAndTime(endDate, endTime);

      // Prepare the item data with CORRECT column names
      const itemData: {
        Title: string;
        Dato: string;
        SlutDato: string;
        AdministratorId: number;
        Placering: string;
        Capacity: number | null;
        M_x00e5_lgruppeId?: number[];
        Online?: {
          Description: string;
          Url: string;
        } | null;
      } = {
        Title: title,
        Dato: finalStartDate.toISOString(),
        SlutDato: finalEndDate.toISOString(),
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

      // Only add Target group if users are selected
      if (selectedUsers.length > 0) {
        itemData.M_x00e5_lgruppeId = selectedUsers.map((u) => u.Id);
      }

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

        // Delete all fields in parallel using Promise.all
        if (existingFields.length > 0) {
          await Promise.all(
            existingFields.map((existingField) =>
              sp.web.lists
                .getByTitle("EventFields")
                .items.getById(existingField.Id)
                .delete()
            )
          );
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
        await Promise.all(
          customFields.map((field) =>
            sp.web.lists.getByTitle("EventFields").items.add({
              Title: field.fieldName,
              EventId: eventId,
              FeltType: field.fieldType,
              Valgmuligheder: field.options
                ? JSON.stringify(field.options)
                : null,
            })
          )
        );
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
    startTime,
    endDate,
    endTime,
    context,
    isOnline,
    selectedLocation,
    maxParticipants,
    onlineLink,
    eventToEdit,
    customFields,
    selectedUsers,
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
        <TextField
          label="Title"
          value={title}
          onChange={onTitleChange}
          required
        />

        <div className={styles.dateStyle}>
          <DatePicker
            label="Fra"
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            placeholder="Vælg start dato"
            ariaLabel="Vælg start dato"
            value={startDate}
            onSelectDate={onStartDateChange}
          />

          <TimePicker
            label="Start tidspunkt"
            dateAnchor={startDate}
            onChange={onStartTimeChange}
          />
        </div>

        <div className={styles.dateStyle}>
          <DatePicker
            label="Til"
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            placeholder="Vælg slut dato"
            ariaLabel="Vælg slut dato"
            value={endDate}
            onSelectDate={onEndDateChange}
            disabled={sameDate}
          />

          <TimePicker
            label="Slut tidspunkt"
            dateAnchor={endDate}
            onChange={onEndTimeChange}
          />
        </div>

        <Checkbox
          label="Samme dag?"
          checked={sameDate}
          onChange={sameDateChange}
        />

        <Stack tokens={{ childrenGap: 10 }}>
          <Label>Målgruppe (valgfrit)</Label>
          <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
            Hvis ingen brugere er valgt, kan alle se eventet. Vælg specifikke
            brugere for at begrænse synligheden.
          </Text>

          <SearchBox
            placeholder="Søg efter bruger..."
            value={userSearchText}
            onChange={(_, newValue) => handleUserSearchChange(newValue)}
          />

          {isSearchingUsers && <Spinner size={SpinnerSize.small} />}

          {userSearchResults.length > 0 && (
            <Stack
              tokens={{ childrenGap: 5 }}
              styles={{
                root: {
                  maxHeight: 200,
                  overflowY: "auto",
                  border: "1px solid #ccc",
                  padding: 5,
                },
              }}
            >
              {userSearchResults.map((user) => (
                <Stack
                  key={user.Id}
                  horizontal
                  horizontalAlign="space-between"
                  styles={{
                    root: {
                      padding: 8,
                      backgroundColor: "white",
                      border: "1px solid #edebe9",
                      cursor: "pointer",
                      ":hover": {
                        backgroundColor: "#f3f2f1",
                      },
                    },
                  }}
                >
                  <Stack>
                    <Text style={{ fontWeight: 600 }}>{user.Title}</Text>
                    {user.Email && (
                      <Text style={{ fontSize: 12, color: "#666" }}>
                        {user.Email}
                      </Text>
                    )}
                  </Stack>
                  <PrimaryButton
                    text="Tilføj"
                    onClick={() => addUserToTargetGroup(user)}
                  />
                </Stack>
              ))}
            </Stack>
          )}

          {selectedUsers.length > 0 && (
            <Stack tokens={{ childrenGap: 5 }}>
              <Label>Valgte brugere:</Label>
              {selectedUsers.map((user) => (
                <Stack
                  key={user.Id}
                  horizontal
                  horizontalAlign="space-between"
                  styles={{
                    root: {
                      padding: 8,
                      backgroundColor: "#f3f2f1",
                      border: "1px solid #edebe9",
                    },
                  }}
                >
                  <Stack>
                    <Text style={{ fontWeight: 600 }}>{user.Title}</Text>
                    {user.Email && (
                      <Text style={{ fontSize: 12, color: "#666" }}>
                        {user.Email}
                      </Text>
                    )}
                  </Stack>
                  <IconButton
                    iconProps={{ iconName: "Delete" }}
                    title="Fjern bruger"
                    onClick={() => removeUserFromTargetGroup(user.Id)}
                  />
                </Stack>
              ))}
            </Stack>
          )}
        </Stack>

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

import * as React from "react";
import { useState, useEffect, useCallback, useMemo } from "react";
import { getSP } from "../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import {
  Panel,
  PanelType,
  PrimaryButton,
  Stack,
  IconButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  SearchBox,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IManageRegistrationsProps {
  isOpen: boolean;
  onDismiss: () => void;
  context: WebPartContext;
  eventId: number;
  eventTitle: string;
}

interface IRegisteredUser {
  Id: number;
  Title: string;
  BrugerId: number;
  Email?: string;
}

const ManageRegistrations: React.FC<IManageRegistrationsProps> = (props) => {
  const { isOpen, onDismiss, context, eventId, eventTitle } = props;

  // State
  const [registeredUsers, setRegisteredUsers] = useState<IRegisteredUser[]>([]);
  const [waitlistUsers, setWaitlistUsers] = useState<IRegisteredUser[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [showAddUser, setShowAddUser] = useState<string>("");
  const [searchResults, setSearchResults] = useState<
    Array<{ Id: number; Title: string; Email: string }>
  >([]);
  const [searchText, setSearchText] = useState("");
  const [isSearching, setIsSearching] = useState(false);

  const loadRegisteredUsers = useCallback(async (): Promise<void> => {
    try {
      setIsLoading(true);
      setError(undefined);
      const sp = getSP(context);

      // Load registered users
      const registered = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`EventId eq ${eventId} and EventType ne 'Waitlist'`)
        .select("Id", "Title", "BrugerId", "EventType")();

      // Load waitlist users
      const waitlist = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`EventId eq ${eventId} and EventType eq 'Waitlist'`)
        .select("Id", "Title", "BrugerId", "EventType")();

      setRegisteredUsers(registered);
      setWaitlistUsers(waitlist);
      setIsLoading(false);
    } catch (error) {
      console.error("Error loading registrations:", error);
      setIsLoading(false);
      setError("Kunne ikke indlæse tilmeldte brugere");
    }
  }, [context, eventId]);

  // Load registered users when panel opens
  useEffect(() => {
    if (isOpen) {
      loadRegisteredUsers().catch(console.error);
    }
  }, [isOpen, loadRegisteredUsers]);

  const removeUser = useCallback(
    async (registrationId: number): Promise<void> => {
      if (!confirm("Er du sikker p\u00e5, at du vil fjerne denne bruger?")) {
        return;
      }

      try {
        const sp = getSP(context);

        await sp.web.lists
          .getByTitle("EventRegistrations")
          .items.getById(registrationId)
          .delete();

        alert("Bruger fjernet!");
        await loadRegisteredUsers();
      } catch (error) {
        console.error("Error removing user:", error);
        alert("Fejl ved fjernelse af bruger. Pr\u00f8v igen.");
      }
    },
    [context, loadRegisteredUsers]
  );

  const searchUsers = useCallback(
    async (searchText: string): Promise<void> => {
      if (!searchText || searchText.length < 2) {
        setSearchResults([]);
        return;
      }

      try {
        setIsSearching(true);
        const sp = getSP(context);

        // Search for users in SharePoint
        const users = await sp.web.siteUsers
          .filter(`substringof('${searchText}', Title)`)
          .top(10)();

        const results = users.map((user) => ({
          Id: user.Id,
          Title: user.Title,
          Email: user.Email || "",
        }));

        setSearchResults(results);
        setIsSearching(false);
      } catch (error) {
        console.error("Error searching users:", error);
        setIsSearching(false);
      }
    },
    [context]
  );

  const addUser = useCallback(
    async (
      userId: number,
      userName: string,
      eventType: string = "Registered"
    ): Promise<void> => {
      try {
        const sp = getSP(context);

        // Check if user is already registered
        const existing = await sp.web.lists
          .getByTitle("EventRegistrations")
          .items.filter(`EventId eq ${eventId} and BrugerId eq ${userId}`)
          .select("Id")();

        if (existing.length > 0) {
          alert("Denne bruger er allerede tilmeldt eventet.");
          return;
        }

        // Add registration
        const registrationKey = `${eventId}_${userId}_${new Date().getTime()}`;

        await sp.web.lists.getByTitle("EventRegistrations").items.add({
          Title: userName,
          EventId: eventId,
          BrugerId: userId,
          EventType: eventType,
          RegistrationKey: registrationKey,
          Submitted: new Date().toISOString(),
        });

        alert("Bruger tilmeldt!");
        setShowAddUser("");
        setSearchText("");
        setSearchResults([]);
        await loadRegisteredUsers();
      } catch (error) {
        console.error("Error adding user:", error);
        alert("Fejl ved tilmelding af bruger. Prøv igen.");
      }
    },
    [context, eventId, loadRegisteredUsers]
  );

  const columns = useMemo((): IColumn[] => {
    return [
      {
        key: "Title",
        name: "Navn",
        fieldName: "Title",
        minWidth: 150,
        maxWidth: 250,
        isResizable: true,
      },
      {
        key: "actions",
        name: "Slet bruger",
        fieldName: "actions",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IRegisteredUser) => {
          return (
            <IconButton
              iconProps={{ iconName: "Delete" }}
              title="Fjern bruger"
              onClick={() => removeUser(item.Id)}
            />
          );
        },
      },
    ];
  }, [removeUser]);

  const handleSearchChange = useCallback(
    (newValue: string | undefined): void => {
      const value = newValue || "";
      setSearchText(value);
      searchUsers(value).catch(console.error);
    },
    [searchUsers]
  );

  const handleToggleAddUser = useCallback((type: string): void => {
    setShowAddUser((prev) => (prev === type ? "" : type));
  }, []);

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={`Administrer tilmeldinger til - ${eventTitle}`}
      closeButtonAriaLabel="Luk"
    >
      <Stack tokens={{ childrenGap: 15 }}>
        {isLoading && <Spinner size={SpinnerSize.large} label="Indlæser..." />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {!isLoading && !error && (
          <>
            {/* Registered Users Section */}
            <Stack horizontal horizontalAlign="space-between">
              <h3>Tilmeldte brugere ({registeredUsers.length})</h3>
              <PrimaryButton
                text="Tilføj bruger"
                iconProps={{ iconName: "Add" }}
                onClick={() => handleToggleAddUser("registered")}
              />
            </Stack>

            {showAddUser === "registered" && (
              <Stack
                tokens={{ childrenGap: 10 }}
                styles={{ root: { padding: 10, backgroundColor: "#f3f2f1" } }}
              >
                <SearchBox
                  placeholder="Søg efter bruger..."
                  value={searchText}
                  onChange={(_, newValue) => handleSearchChange(newValue)}
                />

                {isSearching && <Spinner size={SpinnerSize.small} />}

                {searchResults.length > 0 && (
                  <Stack tokens={{ childrenGap: 5 }}>
                    {searchResults.map((user) => (
                      <Stack
                        key={user.Id}
                        horizontal
                        horizontalAlign="space-between"
                        styles={{
                          root: {
                            padding: 8,
                            backgroundColor: "white",
                            border: "1px solid #ccc",
                          },
                        }}
                      >
                        <Stack>
                          <span style={{ fontWeight: 600 }}>{user.Title}</span>
                          {user.Email && (
                            <span style={{ fontSize: 12, color: "#666" }}>
                              {user.Email}
                            </span>
                          )}
                        </Stack>
                        <PrimaryButton
                          text="Tilføj"
                          onClick={() =>
                            addUser(user.Id, user.Title, "Registered")
                          }
                        />
                      </Stack>
                    ))}
                  </Stack>
                )}
              </Stack>
            )}

            {registeredUsers.length > 0 ? (
              <DetailsList
                items={registeredUsers}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
              />
            ) : (
              <MessageBar messageBarType={MessageBarType.info}>
                Ingen tilmeldte brugere
              </MessageBar>
            )}

            {/* Waitlist Section */}
            <Stack
              horizontal
              horizontalAlign="space-between"
              styles={{ root: { marginTop: 20 } }}
            >
              <h3>Venteliste ({waitlistUsers.length})</h3>
              <PrimaryButton
                text="Tilføj bruger"
                iconProps={{ iconName: "Add" }}
                onClick={() => handleToggleAddUser("waitlist")}
              />
            </Stack>

            {showAddUser === "waitlist" && (
              <Stack
                tokens={{ childrenGap: 10 }}
                styles={{ root: { padding: 10, backgroundColor: "#f3f2f1" } }}
              >
                <SearchBox
                  placeholder="Søg efter bruger..."
                  value={searchText}
                  onChange={(_, newValue) => handleSearchChange(newValue)}
                />

                {isSearching && <Spinner size={SpinnerSize.small} />}

                {searchResults.length > 0 && (
                  <Stack tokens={{ childrenGap: 5 }}>
                    {searchResults.map((user) => (
                      <Stack
                        key={user.Id}
                        horizontal
                        horizontalAlign="space-between"
                        styles={{
                          root: {
                            padding: 8,
                            backgroundColor: "white",
                            border: "1px solid #ccc",
                          },
                        }}
                      >
                        <Stack>
                          <span style={{ fontWeight: 600 }}>{user.Title}</span>
                          {user.Email && (
                            <span style={{ fontSize: 12, color: "#666" }}>
                              {user.Email}
                            </span>
                          )}
                        </Stack>
                        <PrimaryButton
                          text="Tilføj"
                          onClick={() =>
                            addUser(user.Id, user.Title, "Waitlist")
                          }
                        />
                      </Stack>
                    ))}
                  </Stack>
                )}
              </Stack>
            )}

            {waitlistUsers.length > 0 ? (
              <DetailsList
                items={waitlistUsers}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
              />
            ) : (
              <MessageBar messageBarType={MessageBarType.info}>
                Ingen brugere på ventelisten
              </MessageBar>
            )}
          </>
        )}
      </Stack>
    </Panel>
  );
};

export default ManageRegistrations;

import * as React from "react";
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

interface IManageRegistrationsState {
  registeredUsers: IRegisteredUser[];
  isLoading: boolean;
  error?: string;
  showAddUser: boolean;
  searchResults: Array<{ Id: number; Title: string; Email: string }>;
  searchText: string;
  isSearching: boolean;
}

export default class ManageRegistrations extends React.Component<
  IManageRegistrationsProps,
  IManageRegistrationsState
> {
  constructor(props: IManageRegistrationsProps) {
    super(props);
    this.state = {
      registeredUsers: [],
      isLoading: false,
      showAddUser: false,
      searchResults: [],
      searchText: "",
      isSearching: false,
    };
  }

  public componentDidMount(): void {
    if (this.props.isOpen) {
      this.loadRegisteredUsers().catch(console.error);
    }
  }

  public componentDidUpdate(prevProps: IManageRegistrationsProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this.loadRegisteredUsers().catch(console.error);
    }
  }

  private loadRegisteredUsers = async (): Promise<void> => {
    try {
      this.setState({ isLoading: true, error: undefined });
      const sp = getSP(this.props.context);

      const registrations = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`EventId eq ${this.props.eventId}`)
        .select("Id", "Title", "BrugerId")();

      this.setState({
        registeredUsers: registrations,
        isLoading: false,
      });
    } catch (error) {
      console.error("Error loading registrations:", error);
      this.setState({
        isLoading: false,
        error: "Kunne ikke indlæse tilmeldte brugere",
      });
    }
  };

  private removeUser = async (registrationId: number): Promise<void> => {
    if (!confirm("Er du sikker på, at du vil fjerne denne bruger?")) {
      return;
    }

    try {
      const sp = getSP(this.props.context);

      await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.getById(registrationId)
        .delete();

      alert("Bruger fjernet!");
      await this.loadRegisteredUsers();
    } catch (error) {
      console.error("Error removing user:", error);
      alert("Fejl ved fjernelse af bruger. Prøv igen.");
    }
  };

  private searchUsers = async (searchText: string): Promise<void> => {
    if (!searchText || searchText.length < 2) {
      this.setState({ searchResults: [] });
      return;
    }

    try {
      this.setState({ isSearching: true });
      const sp = getSP(this.props.context);

      // Search for users in SharePoint
      const users = await sp.web.siteUsers
        .filter(`substringof('${searchText}', Title)`)
        .top(10)();

      const searchResults = users.map((user) => ({
        Id: user.Id,
        Title: user.Title,
        Email: user.Email || "",
      }));

      this.setState({ searchResults, isSearching: false });
    } catch (error) {
      console.error("Error searching users:", error);
      this.setState({ isSearching: false });
    }
  };

  private addUser = async (userId: number, userName: string): Promise<void> => {
    try {
      const sp = getSP(this.props.context);

      // Check if user is already registered
      const existing = await sp.web.lists
        .getByTitle("EventRegistrations")
        .items.filter(`EventId eq ${this.props.eventId} and BrugerId eq ${userId}`)
        .select("Id")();

      if (existing.length > 0) {
        alert("Denne bruger er allerede tilmeldt eventet.");
        return;
      }

      // Add registration
      const registrationKey = `${this.props.eventId}_${userId}_${new Date().getTime()}`;

      await sp.web.lists.getByTitle("EventRegistrations").items.add({
        Title: userName,
        EventId: this.props.eventId,
        BrugerId: userId,
        RegistrationKey: registrationKey,
        Submitted: new Date().toISOString(),
      });

      alert("Bruger tilmeldt!");
      this.setState({ showAddUser: false, searchText: "", searchResults: [] });
      await this.loadRegisteredUsers();
    } catch (error) {
      console.error("Error adding user:", error);
      alert("Fejl ved tilmelding af bruger. Prøv igen.");
    }
  };

  private getColumns = (): IColumn[] => {
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
              onClick={() => this.removeUser(item.Id)}
            />
          );
        },
      },
    ];
  };

  public render(): React.ReactElement {
    const { isOpen, onDismiss, eventTitle } = this.props;
    const { registeredUsers, isLoading, error, showAddUser, searchResults, searchText, isSearching } = this.state;

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
              <Stack horizontal horizontalAlign="space-between">
                <h3>Tilmeldte brugere ({registeredUsers.length})</h3>
                <PrimaryButton
                  text="Tilføj bruger"
                  iconProps={{ iconName: "Add" }}
                  onClick={() => this.setState({ showAddUser: !showAddUser })}
                />
              </Stack>

              {showAddUser && (
                <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 10, backgroundColor: "#f3f2f1" } }}>
                  <SearchBox
                    placeholder="Søg efter bruger..."
                    value={searchText}
                    onChange={(_, newValue) => {
                      this.setState({ searchText: newValue || "" });
                      this.searchUsers(newValue || "").catch(console.error);
                    }}
                  />

                  {isSearching && <Spinner size={SpinnerSize.small} />}

                  {searchResults.length > 0 && (
                    <Stack tokens={{ childrenGap: 5 }}>
                      {searchResults.map((user) => (
                        <Stack
                          key={user.Id}
                          horizontal
                          horizontalAlign="space-between"
                          styles={{ root: { padding: 8, backgroundColor: "white", border: "1px solid #ccc" } }}
                        >
                          <Stack>
                            <span style={{ fontWeight: 600 }}>{user.Title}</span>
                            {user.Email && <span style={{ fontSize: 12, color: "#666" }}>{user.Email}</span>}
                          </Stack>
                          <PrimaryButton
                            text="Tilføj"
                            onClick={() => this.addUser(user.Id, user.Title)}
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
                  columns={this.getColumns()}
                  selectionMode={SelectionMode.none}
                  layoutMode={DetailsListLayoutMode.justified}
                  isHeaderVisible={true}
                />
              ) : (
                <MessageBar messageBarType={MessageBarType.info}>
                  Ingen tilmeldte brugere
                </MessageBar>
              )}
            </>
          )}
        </Stack>
      </Panel>
    );
  }
}
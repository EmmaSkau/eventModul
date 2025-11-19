import * as React from "react";
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

interface IRegisterForEventsState {
  fields: IEventField[];
  fieldValues: { [key: number]: string | boolean };
  isLoading: boolean;
  isSaving: boolean;
  error?: string;
  success?: string;
}

export default class RegisterForEvents extends React.Component<
  IRegisterForEventsProps,
  IRegisterForEventsState
> {
  constructor(props: IRegisterForEventsProps) {
    super(props);
    this.state = {
      fields: [],
      fieldValues: {},
      isLoading: false,
      isSaving: false,
    };
  }

  public componentDidMount(): void {
    if (this.props.isOpen) {
      this.loadEventFields().catch(console.error);
    }
  }

  public componentDidUpdate(prevProps: IRegisterForEventsProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      this.loadEventFields().catch(console.error);
    }
  }

  private loadEventFields = async (): Promise<void> => {
    try {
      this.setState({ isLoading: true, error: undefined });
      const sp = getSP(this.props.context);
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

      const fields: IEventField[] = allItems.filter(
        (item) => item.EventId === this.props.eventId
      );

      this.setState({
        fields,
        isLoading: false,
      });
    } catch (error) {
      console.error("Error loading event fields:", error);
      this.setState({
        isLoading: false,
        error: "Kunne ikke indlæse tilmeldingsfelter",
      });
    }
  };

  private onFieldChange = (fieldId: number, value: string | boolean): void => {
    this.setState((prevState) => ({
      fieldValues: {
        ...prevState.fieldValues,
        [fieldId]: value,
      },
    }));
  };

  private validateForm = (): boolean => {
    const { fields, fieldValues } = this.state;

    for (const field of fields) {
      if (field.P_x00e5_kr_x00e6_vet) {
        const value = fieldValues[field.Id];
        if (value === undefined || value === "" || value === false) {
          this.setState({
            error: `Feltet "${field.Title}" er påkrævet`,
          });
          return false;
        }
      }
    }

    return true;
  };

  private handleSubmit = async (): Promise<void> => {
    if (!this.validateForm()) {
      return;
    }

    try {
      this.setState({ isSaving: true, error: undefined });
      const sp = getSP(this.props.context);

      // Generate a unique registration key
      const registrationKey = `${this.props.eventId}_${
        this.props.context.pageContext.user.loginName
      }_${new Date().getTime()}`;

      const currentUser = await sp.web.currentUser();

      for (const field of this.state.fields) {
        const value = this.state.fieldValues[field.Id];
        if (value !== undefined) {
          const itemData = {
            Title: this.props.eventTitle,
            EventId: this.props.eventId, 
            BrugerId: currentUser.Id, 
            FieldName: field.Title,
            FieldType: field.FeltType,
            FieldValue: String(value),
            RegistrationKey: registrationKey,
            Submitted: new Date().toISOString(),
          };

          await sp.web.lists.getByTitle("EventRegistrations").items.add(itemData);
        }
      }

      this.setState({
        isSaving: false,
        success: "Du er nu tilmeldt eventet!",
        fieldValues: {},
      });

      // Close panel after 2 seconds
      setTimeout(() => {
        this.props.onDismiss();
      }, 2000);
    } catch (error) {
      console.error("Error submitting registration:", error);
      this.setState({
        isSaving: false,
        error: "Kunne ikke gemme tilmeldingen. Prøv igen.",
      });
    }
  };

  private renderField = (field: IEventField): JSX.Element => {
    const value = this.state.fieldValues[field.Id] || "";

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
              this.onFieldChange(field.Id, newValue || "")
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
              this.onFieldChange(field.Id, (option?.key as string) || "")
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
            onChange={(_, checked) => this.onFieldChange(field.Id, !!checked)}
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
              this.onFieldChange(field.Id, newValue || "")
            }
          />
        );
    }
  };

  public render(): React.ReactElement<IRegisterForEventsProps> {
    const { isOpen, onDismiss, eventTitle } = this.props;
    const { fields, isLoading, isSaving, error, success } = this.state;

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
              <div key={field.Id}>{this.renderField(field)}</div>
            ))}

            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton
                text={isSaving ? "Gemmer..." : "Tilmeld"}
                onClick={this.handleSubmit}
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
  }
}

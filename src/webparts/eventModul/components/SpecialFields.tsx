import * as React from "react";
import {
  TextField,
  Dropdown,
  IDropdownOption,
  Stack,
  PrimaryButton,
  DefaultButton,
} from "@fluentui/react";

export interface ICustomField {
  id: string;
  fieldName: string;
  fieldType: "text" | "multipleChoice";
  options?: string[];
}

interface IAddFieldDialogProps {
  onAddField: (field: ICustomField) => void;
  onCancel: () => void;
}

interface IAddFieldDialogState {
  newFieldName: string;
  newFieldType: "text" | "multipleChoice";
  newFieldOptions: string;
}

export default class AddFieldDialog extends React.Component<
  IAddFieldDialogProps,
  IAddFieldDialogState
> {
  constructor(props: IAddFieldDialogProps) {
    super(props);

    this.state = {
      newFieldName: "",
      newFieldType: "text",
      newFieldOptions: "",
    };
  }

  private onFieldNameChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ newFieldName: newValue || "" });
  };

  private onFieldTypeChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ newFieldType: option.key as "text" | "multipleChoice" });
    }
  };

  private onFieldOptionsChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ newFieldOptions: newValue || "" });
  };

  private handleAddField = (): void => {
    const { newFieldName, newFieldType, newFieldOptions } = this.state;

    // Validation
    if (!newFieldName.trim()) {
      alert("Indtast venligst et feltnavn");
      return;
    }

    if (newFieldType === "multipleChoice" && !newFieldOptions.trim()) {
      alert("Indtast venligst valgmuligheder for flervalg");
      return;
    }

    // Create the field object
    const field: ICustomField = {
      id: Date.now().toString(),
      fieldName: newFieldName.trim(),
      fieldType: newFieldType,
      options:
        newFieldType === "multipleChoice"
          ? newFieldOptions.split(",").map((opt) => opt.trim()).filter((opt) => opt)
          : undefined,
    };

    // Call the parent callback
    this.props.onAddField(field);

    // Reset the form
    this.setState({
      newFieldName: "",
      newFieldType: "text",
      newFieldOptions: "",
    });
  };

  private handleCancel = (): void => {
    // Reset the form
    this.setState({
      newFieldName: "",
      newFieldType: "text",
      newFieldOptions: "",
    });

    // Call parent cancel
    this.props.onCancel();
  };

  public render(): React.ReactElement {
    const { newFieldName, newFieldType, newFieldOptions } = this.state;

    const fieldTypeOptions: IDropdownOption[] = [
      { key: "text", text: "Tekstfelt" },
      { key: "multipleChoice", text: "Flervalg" },
    ];

    return (
      <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: '10px 0' } }}>
        <TextField
          label="Feltnavn"
          placeholder="fx T-shirt størrelse"
          value={newFieldName}
          onChange={this.onFieldNameChange}
          required
        />

        <Dropdown
          label="Felttype"
          options={fieldTypeOptions}
          selectedKey={newFieldType}
          onChange={this.onFieldTypeChange}
          required
        />

        {newFieldType === "multipleChoice" && (
          <TextField
            label="Valgmuligheder (adskilt med komma)"
            placeholder="Vegan, Vegetar, Ingen restriktioner"
            multiline
            rows={3}
            value={newFieldOptions}
            onChange={this.onFieldOptionsChange}
            required
          />
        )}

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton text="Tilføj" onClick={this.handleAddField} />
          <DefaultButton text="Annuller" onClick={this.handleCancel} />
        </Stack>
      </Stack>
    );
  }
}

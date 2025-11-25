import * as React from "react";
import { useState, useCallback } from "react";
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

const AddFieldDialog: React.FC<IAddFieldDialogProps> = ({
  onAddField,
  onCancel,
}) => {
  // State
  const [newFieldName, setNewFieldName] = useState("");
  const [newFieldType, setNewFieldType] = useState<"text" | "multipleChoice">(
    "text"
  );
  const [newFieldOptions, setNewFieldOptions] = useState("");

  // Event handlers
  const onFieldNameChange = useCallback(
    (
      event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setNewFieldName(newValue || "");
    },
    []
  );

  const onFieldTypeChange = useCallback(
    (
      event: React.FormEvent<HTMLDivElement>,
      option?: IDropdownOption
    ): void => {
      if (option) {
        setNewFieldType(option.key as "text" | "multipleChoice");
      }
    },
    []
  );

  const onFieldOptionsChange = useCallback(
    (
      event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setNewFieldOptions(newValue || "");
    },
    []
  );

  const handleAddField = useCallback((): void => {
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
          ? newFieldOptions
              .split(",")
              .map((opt) => opt.trim())
              .filter((opt) => opt)
          : undefined,
    };

    // Call the parent callback
    onAddField(field);

    // Reset the form
    setNewFieldName("");
    setNewFieldType("text");
    setNewFieldOptions("");
  }, [newFieldName, newFieldType, newFieldOptions, onAddField]);

  const handleCancel = useCallback((): void => {
    // Reset the form
    setNewFieldName("");
    setNewFieldType("text");
    setNewFieldOptions("");

    // Call parent cancel
    onCancel();
  }, [onCancel]);

  const fieldTypeOptions: IDropdownOption[] = [
    { key: "text", text: "Tekstfelt" },
    { key: "multipleChoice", text: "Flervalg" },
  ];

  return (
    <Stack
      tokens={{ childrenGap: 15 }}
      styles={{ root: { padding: "10px 0" } }}
    >
      <TextField
        label="Feltnavn"
        placeholder="fx T-shirt størrelse"
        value={newFieldName}
        onChange={onFieldNameChange}
        required
      />

      <Dropdown
        label="Felttype"
        options={fieldTypeOptions}
        selectedKey={newFieldType}
        onChange={onFieldTypeChange}
        required
      />

      {newFieldType === "multipleChoice" && (
        <TextField
          label="Valgmuligheder (adskilt med komma)"
          placeholder="Vegan, Vegetar, Ingen restriktioner"
          multiline
          rows={3}
          value={newFieldOptions}
          onChange={onFieldOptionsChange}
          required
        />
      )}

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton text="Tilføj" onClick={handleAddField} />
        <DefaultButton text="Annuller" onClick={handleCancel} />
      </Stack>
    </Stack>
  );
};

export default AddFieldDialog;

import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import styles from "../EventModul.module.scss";
import {
  IDropdownOption,
  DatePicker,
  DayOfWeek,
  Dropdown,
  PrimaryButton,
} from "@fluentui/react";
import { formatDate } from "./formatDate";
import { getSP } from "../../../../pnpConfig";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface ILocationItem {
  Placering: string;
}

interface IFilterUtilityProps {
  context: WebPartContext;
  startDate?: Date | undefined;
  endDate?: Date | undefined;
  selectedLocation?: string | undefined;
  locationOptions?: IDropdownOption[];
  isLoadingLocations?: boolean;
  onLocationChange?: (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => void;
  onStartDateSelect?: (date: Date | undefined | undefined) => void;
  onEndDateSelect?: (date: Date | undefined | undefined) => void;
  resetFilters?: () => void;
}

interface IUseLocationFilterReturn {
  startDate: Date | undefined;
  endDate: Date | undefined;
  selectedLocation: string | undefined;
  locationOptions: IDropdownOption[];
  isLoadingLocations: boolean;
  onLocationChange: (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => void;
  onStartDateSelect: (date: Date | undefined | undefined) => void;
  onEndDateSelect: (date: Date | undefined | undefined) => void;
  setStartDate: React.Dispatch<React.SetStateAction<Date | undefined>>;
  setEndDate: React.Dispatch<React.SetStateAction<Date | undefined>>;
  setSelectedLocation: React.Dispatch<React.SetStateAction<string | undefined>>;
  loadLocationsFromSharePoint: () => Promise<void>;
  resetFilters: () => void;
}

// Custom hook for filtering logic
export const useLocationFilter = (
  context: WebPartContext
): IUseLocationFilterReturn => {
  // State
  const [startDate, setStartDate] = useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = useState<Date | undefined>(undefined);
  const [selectedLocation, setSelectedLocation] = useState<string | undefined>(
    undefined
  );
  const [locationOptions, setLocationOptions] = useState<IDropdownOption[]>([]);
  const [isLoadingLocations, setIsLoadingLocations] = useState(false);

  const loadLocationsFromSharePoint = useCallback(async (): Promise<void> => {
    try {
      setIsLoadingLocations(true);
      const sp = getSP(context);

      // Get all items with Placering field (Location field can't be filtered in query)
      const items: ILocationItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.select("Placering")();

      console.log("Loaded items from SharePoint:", items);

      const locations = items
        .map((item: ILocationItem) => {
          if (!item.Placering) return null;
          try {
            const parsed = JSON.parse(item.Placering);
            return parsed.DisplayName || item.Placering;
          } catch {
            return item.Placering;
          }
        })
        .filter((location) => location);
      const uniqueLocations: string[] = [];
      const seen: { [key: string]: boolean } = {};
      for (const location of locations) {
        if (!seen[location]) {
          seen[location] = true;
          uniqueLocations.push(location);
        }
      }

      const options: IDropdownOption[] = [
        { key: "all", text: "Alle lokationer" },
        ...uniqueLocations.map((location: string) => ({
          key: location,
          text: location,
        })),
      ];

      setLocationOptions(options);
      setIsLoadingLocations(false);
    } catch (error) {
      console.error("Error loading locations from SharePoint:", error);
      setIsLoadingLocations(false);
      setLocationOptions([{ key: "all", text: "Alle lokationer" }]);
    }
  }, [context]);

  // Load locations on mount
  useEffect(() => {
    loadLocationsFromSharePoint().catch((error) => {
      console.error("Error in componentDidMount:", error);
    });
  }, [loadLocationsFromSharePoint]);

  const onLocationChange = useCallback(
    (
      event: React.FormEvent<HTMLDivElement>,
      option?: IDropdownOption
    ): void => {
      if (option) {
        setSelectedLocation(option.key as string);
      }
    },
    []
  );

  const onStartDateSelect = useCallback(
    (date: Date | null | undefined): void => {
      setStartDate(date || undefined);
    },
    []
  );

  const onEndDateSelect = useCallback((date: Date | null | undefined): void => {
    setEndDate(date || undefined);
  }, []);

  const resetFilters = useCallback((): void => {
    setStartDate(undefined);
    setEndDate(undefined);
    setSelectedLocation(undefined);
  }, []);

  return {
    startDate,
    endDate,
    selectedLocation,
    locationOptions,
    isLoadingLocations,
    onLocationChange,
    onStartDateSelect,
    onEndDateSelect,
    setStartDate,
    setEndDate,
    setSelectedLocation,
    loadLocationsFromSharePoint,
    resetFilters,
  };
};

// UI Component
const FilterUtility: React.FC<IFilterUtilityProps> = ({
  context,
  startDate,
  endDate,
  selectedLocation,
  locationOptions,
  isLoadingLocations,
  onLocationChange,
  onStartDateSelect,
  onEndDateSelect,
  resetFilters,
}) => {
  // If no props provided, use the hook (for backward compatibility)
  const hookValues = useLocationFilter(context);

  // Use props if provided, otherwise fall back to hook values
  const finalStartDate =
    startDate !== undefined ? startDate : hookValues.startDate;
  const finalEndDate = endDate !== undefined ? endDate : hookValues.endDate;
  const finalSelectedLocation =
    selectedLocation !== undefined
      ? selectedLocation
      : hookValues.selectedLocation;
  const finalLocationOptions = locationOptions || hookValues.locationOptions;
  const finalIsLoadingLocations =
    isLoadingLocations !== undefined
      ? isLoadingLocations
      : hookValues.isLoadingLocations;
  const finalOnLocationChange = onLocationChange || hookValues.onLocationChange;
  const finalOnStartDateSelect =
    onStartDateSelect || hookValues.onStartDateSelect;
  const finalOnEndDateSelect = onEndDateSelect || hookValues.onEndDateSelect;
  const finalResetFilters = resetFilters || hookValues.resetFilters;

  return (
    <section className={styles.filters}>
      <DatePicker
        label="Fra"
        firstDayOfWeek={DayOfWeek.Monday}
        placeholder="Vælg start dato..."
        ariaLabel="Vælg start dato"
        value={finalStartDate}
        onSelectDate={finalOnStartDateSelect}
        formatDate={formatDate}
      />

      <DatePicker
        label="Til"
        firstDayOfWeek={DayOfWeek.Monday}
        placeholder="Vælg slut dato..."
        ariaLabel="Vælg slut dato"
        value={finalEndDate}
        onSelectDate={finalOnEndDateSelect}
        formatDate={formatDate}
      />

      <Dropdown
        label="Lokation"
        placeholder={
          finalIsLoadingLocations
            ? "Indlæser lokationer..."
            : "Vælg lokation..."
        }
        options={finalLocationOptions}
        selectedKey={finalSelectedLocation}
        onChange={finalOnLocationChange}
        disabled={finalIsLoadingLocations}
      />

      <PrimaryButton
        text="Ryd filtre"
        onClick={finalResetFilters}
      />
    </section>
  );
};

export default FilterUtility;

import * as React from "react";
import styles from "./EventModul.module.scss";
import type { IEventModulProps } from "./IEventModulProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  DatePicker,
  DayOfWeek,
  Dropdown,
  IDropdownOption,
  Toggle,
  DefaultButton,
} from "@fluentui/react";
import { formatDate } from "./Utility/formatDate";
import { getSP } from "../../../pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import ListView from "./ListView";

interface ILocationItem {
  Placering: string;
}

interface IEventModulState {
  isTrue: boolean;
  startDate?: Date;
  endDate?: Date;
  selectedLocation?: string;
  locationOptions: IDropdownOption[];
  isLoadingLocations: boolean;
  registered?: boolean;
  cancelledEvents?: boolean;
  waitlisted?: boolean;
}

export default class EventModul extends React.Component<
  IEventModulProps,
  IEventModulState
> {
  constructor(props: IEventModulProps) {
    super(props);
    this.state = {
      isTrue: false,
      startDate: undefined,
      endDate: undefined,
      selectedLocation: undefined,
      locationOptions: [],
      isLoadingLocations: false,
      registered: false,
      cancelledEvents: false,
      waitlisted: false,
    };
  }

  public componentDidMount(): void {
    this.loadLocationsFromSharePoint().catch((error) => {
      console.error("Error in componentDidMount:", error);
    });
  }

  private loadLocationsFromSharePoint = async (): Promise<void> => {
    try {
      this.setState({ isLoadingLocations: true });
      const sp = getSP(this.props.context);

      // Get all items with Placering field (Location field can't be filtered in query)
      const items: ILocationItem[] = await sp.web.lists
        .getByTitle("EventDB")
        .items.select("Placering")();

      console.log("Loaded items from SharePoint:", items);

      const locations = items
        .map((item: ILocationItem) => item.Placering)
        .filter((location: string) => location); // Filter out null/empty in JavaScript
      const uniqueLocations: string[] = [];
      const seen: { [key: string]: boolean } = {};
      for (const location of locations) {
        if (!seen[location]) {
          seen[location] = true;
          uniqueLocations.push(location);
        }
      }

      const locationOptions: IDropdownOption[] = [
        { key: "all", text: "Alle lokationer" },
        ...uniqueLocations.map((location: string) => ({
          key: location,
          text: location,
        })),
      ];

      this.setState({
        locationOptions,
        isLoadingLocations: false,
      });
    } catch (error) {
      console.error("Error loading locations from SharePoint:", error);
      this.setState({
        isLoadingLocations: false,
        locationOptions: [{ key: "all", text: "Alle lokationer" }],
      });
    }
  };

  private onLocationChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ selectedLocation: option.key as string });
    }
  };

  private onStartDateSelect = (date: Date | null | undefined): void => {
    this.setState({ startDate: date || undefined });
  };

  private onEndDateSelect = (date: Date | null | undefined): void => {
    this.setState({ endDate: date || undefined });
  };

  private resetFilters = (): void => {
    this.setState({
      startDate: undefined,
      endDate: undefined,
      selectedLocation: undefined,
      registered: false,
      cancelledEvents: false,
      waitlisted: false,
    });
  };

  public render(): React.ReactElement<IEventModulProps> {
    const { hasTeamsContext, userDisplayName } = this.props;
    const {
      startDate,
      endDate,
      selectedLocation,
      locationOptions,
      isLoadingLocations,
    } = this.state;

    return (
      <section
        className={`${styles.eventModul} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <h2>{escape(userDisplayName)}s Events</h2>
          <p>Her kan du se alle dine events og fremtidige events</p>
        </div>

        <section className={styles.filters}>
          <DatePicker
            label="Fra"
            firstDayOfWeek={DayOfWeek.Monday}
            placeholder="Vælg start dato..."
            ariaLabel="Vælg start dato"
            value={startDate}
            onSelectDate={this.onStartDateSelect}
            formatDate={formatDate}
          />

          <DatePicker
            label="Til"
            firstDayOfWeek={DayOfWeek.Monday}
            placeholder="Vælg slut dato..."
            ariaLabel="Vælg slut dato"
            value={endDate}
            onSelectDate={this.onEndDateSelect}
            formatDate={formatDate}
          />

          <Dropdown
            label="Lokation"
            placeholder={
              isLoadingLocations ? "Indlæser lokationer..." : "Vælg lokation..."
            }
            options={locationOptions}
            selectedKey={selectedLocation}
            onChange={this.onLocationChange}
            disabled={isLoadingLocations}
          />
        </section>
        <section className={styles.filterToggle}>
          <Toggle label="Tilmeldt" checked={this.state.registered} />

          <Toggle label="Afmeldt" checked={this.state.cancelledEvents} />

          <Toggle label="Venteliste" checked={this.state.waitlisted} />

          <DefaultButton
            className={styles.restFilter}
            text="Ryd filtre"
            onClick={this.resetFilters}
          />
        </section>

        <h2>Fremtidige events:</h2>
        <ListView context={this.props.context} />
      </section>
    );
  }
}

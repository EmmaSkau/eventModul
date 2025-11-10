import * as React from "react";
import { DatePicker, DayOfWeek, PrimaryButton } from "@fluentui/react";

interface ICreateEventProps {
  onClose?: () => void;
}

export default class CreateEvent extends React.Component<ICreateEventProps> {
  public render(): React.ReactElement {
    return (
      <section>
        <h1>Opret ny event</h1>
        <div>
          <DatePicker
            label="Fra"
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            placeholder="Select start date"
            ariaLabel="Select start date"
          />
          <DatePicker
            label="Til"
            firstDayOfWeek={DayOfWeek.Monday}
            showWeekNumbers={true}
            placeholder="Select end date"
            ariaLabel="Select end date"
          />
        </div>

        <PrimaryButton 
            text="Afslut og gem event" 
            onClick={this.props.onClose}
        />
      </section>
    );
  }
}

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IListViewProps {
  context: WebPartContext;
  startDate?: Date;
  endDate?: Date;
  selectedLocation?: string;
  registered?: boolean;
  cancelledEvents?: boolean;
  waitlisted?: boolean;
  showPastEvents?: boolean;
}
import { IEventItem } from "./IEventItem";

export interface IListViewState {
  events: IEventItem[];
  isLoading: boolean;
  error?: string;
  selectedEventId?: number;
  selectedEventTitle?: string;
}
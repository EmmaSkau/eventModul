import { IEventItem } from "./IEventItem";

export const formatDate = (date: Date | string | number | undefined): string => {
  let newDate: Date = new Date();
  if (date instanceof Date && !isNaN(date.getTime())) {
    newDate = date;
  } else if (date !== null && date !== undefined && typeof date !== 'object') {
    newDate = new Date(date);
  }
  return newDate.getDate() + "." + (newDate.getMonth()+1) + "." + newDate.getFullYear();
};

// Filters out events that are before today
export const filterFutureEvents = (items: IEventItem[]): IEventItem[] => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return items.filter((item) => {
    if (!item.Dato) return false;
    const eventDate = new Date(item.Dato);
    eventDate.setHours(0, 0, 0, 0);
    return eventDate >= today;
  });
};

// Sorts events by date (earliest first)
export const sortEventsByDate = (items: IEventItem[]): IEventItem[] => {
  return [...items].sort((a, b) => {
    if (!a.Dato) return 1;
    if (!b.Dato) return -1;
    return new Date(a.Dato).getTime() - new Date(b.Dato).getTime();
  });
};

// Filters future events and sorts them by date
export const getFutureEventsSorted = (items: IEventItem[]): IEventItem[] => {
  const futureEvents = filterFutureEvents(items);
  return sortEventsByDate(futureEvents);
};
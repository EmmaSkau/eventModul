export interface IEventItem {
  Id: number;
  Title: string;
  Dato?: string; // StartDato internal name is "Dato"
  SlutDato?: string;
  Administrator?: {
    Title: string;
  };
  Placering?: string;
  targetGroup?: string;
  Beskrivelse?: string;
  TilfoejEkstraInfo?: string;
  Capacity?: number;
}
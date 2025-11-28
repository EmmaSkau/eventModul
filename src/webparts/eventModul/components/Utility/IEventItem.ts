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
  M_x00e5_lgruppeId?: number[];
  Beskrivelse?: string;
  TilfoejEkstraInfo?: string;
  Capacity?: number;
  Online?: {
    Description: string;
    Url: string;
  };
}
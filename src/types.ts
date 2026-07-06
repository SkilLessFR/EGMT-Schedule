export type ShiftCode = 'MID' | 'A' | 'M' | 'N' | 'OFF' | 'H8' | string;

export interface ShiftEvent {
  id: string;
  date: Date;
  isoDate: string;
  shift: ShiftCode;
  notes?: string;
  hours?: string;
}

export interface RosterData {
  month: number;
  year: number;
  employees: string[];
  rows: Record<string, Record<string, string>>;
  dateColumns: { index: number; date: Date; isoDate: string }[];
  fileName: string;
}

export interface ShiftColor {
  bg: string;
  text: string;
  border: string;
}

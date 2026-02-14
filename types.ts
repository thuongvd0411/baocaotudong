
export interface ESDMItem {
  skill: string;
  level0?: string;
  level1: string;
  level2: string;
  level3: string;
  level4: string;
}

export interface StudentInfo {
  name: string;
  dob: string;
  evalDate: string;
  age: string;
  gender: 'Nam' | 'Nữ';
  studentId: string;
}

export interface ESDMResult {
  table: ESDMItem[];
  percents: Record<string, number>;
  percentsOld?: Record<string, number>;
  summary: string;
}

export enum ProcessingStatus {
  IDLE = 'IDLE',
  LOADING = 'LOADING',
  SUCCESS = 'SUCCESS',
  ERROR = 'ERROR'
}

export interface SectionAssignment {
  itemID:number;
  SectionNumber: number;
  SectionName: string;
  POwnerID: number;
  POwnerEmail: string;
  POwnerDisplayName: string,
  SOwnerID: number;
  SOwnerEmail: string;
  SOwnerDisplayName: string,
  Contributors: Contributor[];
  DeadLineDate: Date;
  DocumentID: string;
  SectionID: string;
}

export interface Contributor {
  ContributorEmail?: string;
  ContributorDisplayName: string;
  ContributorID: number;
}

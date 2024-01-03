export interface SectionAssignment {
  SectionNumber: number;
  SectionName: string;
  POwnerID: number;
  POwnerEmail: string;
  POwneDisplayName: string,
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

export interface SectionAssignment {
  SectionNumber: number;
  SectionName: string;
  POwnerID: number;
  POwnerEmail: string;
  POwneDisplayName: string,
  SOwnerID: number;
  SOwnerEmail: string;
  SOwnerDisplayName: string,
  Contributor: Contributors[];
  DeadLineDate: Date;
  DocumentID: string;
  SectionID: string;
}

export interface Contributors {
  Email: string;
  ContributorID: number;
}

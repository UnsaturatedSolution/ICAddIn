export interface SectionAssignment {
  SectionNumber: number;
  POwnerID: number;
  POwnerEmail: string;
  SOwnerID: number;
  SOwnerEmail: string;
  Contributor: Contributors[];
  DeadLineDate: Date;
  DocumentID:string;
}

export interface Contributors {
  Email: string;
  ContributorID: number;
}

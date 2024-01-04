import { DialogType, IColumn } from "@fluentui/react"

export const gridFormDoneColumn:IColumn[]=[
  {
    key: 'Done',
    name: 'Done',
    fieldName: 'Done',
    minWidth: 50,
    maxWidth: 50,
  }
];
export const gridFormManageColumn:IColumn[]=[
  {
    key: 'ManageContributors',
    name: 'Manage Contributors',
    fieldName: 'ManageContributors',
    minWidth: 75,
    maxWidth: 75,
  }
];
export const gridFormColumns: IColumn[] = [
    {
      key: 'SNo',
      name: 'S/N',
      fieldName: 'SNo',
      minWidth: 30,
      maxWidth: 30,
    },
    {
      key: 'Section',
      name: 'Section',
      fieldName: 'Section',
      minWidth: 75,
      maxWidth: 75,
    },
    {
      key: 'Primary',
      name: 'Primary',
      fieldName: 'Primary',
      minWidth: 75,
      maxWidth: 75,
    },
    {
      key: 'Secondary',
      name: 'Secondary',
      fieldName: 'Secondary',
      minWidth: 75,
      maxWidth: 75,
    },
    {
      key: 'TargetDate',
      name: 'Target Date',
      fieldName: 'TargetDate',
      minWidth: 100,
      maxWidth: 100,
    },
    {
      key: 'Contributors',
      name: 'Contributors',
      fieldName: 'Contributors',
      minWidth: 75,
      maxWidth: 75,
    },
    
]
export const contributorFormColumns: IColumn[] = [
    {
      key: 'Contributor',
      name: 'Contributor',
      fieldName: 'Contributor',
      minWidth: 200,
      maxWidth: 200,
    },
    {
      key: 'Action',
      name: 'Action',
      fieldName: 'Action',
      minWidth: 50,
      maxWidth: 50,
    },
]
export const dialogContentProps = {
    type: DialogType.normal,
    title: 'Missing Subject',
    closeButtonAriaLabel: 'Close',
    subText: 'Do you want to send this message without a subject?',
  };
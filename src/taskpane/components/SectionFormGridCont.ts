import { DialogType, IColumn } from "@fluentui/react"

export const gridFormColumns: IColumn[] = [
    {
      key: 'SNo',
      name: 'S/N',
      fieldName: 'SNo',
      minWidth: 50,
      maxWidth: 50,
    },
    {
      key: 'Section',
      name: 'Section',
      fieldName: 'Section',
      minWidth: 100,
      maxWidth: 100,
    },
    {
      key: 'Primary',
      name: 'Primary',
      fieldName: 'Primary',
      minWidth: 100,
      maxWidth: 100,
    },
    {
      key: 'Secondary',
      name: 'Secondary',
      fieldName: 'Secondary',
      minWidth: 100,
      maxWidth: 100,
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
      minWidth: 100,
      maxWidth: 100,
    },
    {
      key: 'ManageContributors',
      name: 'Manage Contributors',
      fieldName: 'ManageContributors',
      minWidth: 100,
      maxWidth: 100,
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
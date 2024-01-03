import React, { Component, ReactElement } from "react";
import {
    Combobox,
    Label,
    Option,
    Persona,
    SelectTabData,
    SelectTabEvent,
    Tab,
    TabList,
    TabValue,
    useId,
} from "@fluentui/react-components";
import type { ComboboxProps } from "@fluentui/react-components";
import { getSearchUser } from "../../helpers/sso-helper";
import { DatePicker, DayOfWeek, DefaultButton, DetailsList, Dialog, DialogFooter, IColumn, Pivot, PivotItem, PrimaryButton, SelectionMode, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import ShowADUserComponent from "./CustomPicker";
import { contributorFormColumns, dialogContentProps, gridFormColumns } from "../../constants/SectionFormGridCont";


export interface IProps extends ComboboxProps {
    // contributors: any[];
    contributorPanelId: number;
    updateParentContributorState: Function;
    isReadOnlyForm: boolean;
    sectionInfo: SectionAssignment;
    allADUsers: any[];
    tempContributors: any;
}

export interface IState {
}


export class ContributorDialog extends Component<IProps, IState> {
    constructor(props: IProps) {
        super(props);
        this.state = {
        };
    }
    public updateContributorState = (newContributor) => {
        let tempCon = [...this.props.tempContributors];
        tempCon.push(newContributor);
        this.props.updateParentContributorState(tempCon);
    }
    public renderGridItems = (item, index: number, column: IColumn) => {
        const fieldValue = item[column.fieldName];
        switch (column.key) {
            case 'Contributor': {
                return <span id={item.ContributorID}>{item.ContributorDisplayName}</span>
                // return <span id={item.ContributorID}>{item.ContributorDisplayName}</span>
            };
            case 'Action': {
                return <DefaultButton
                    style={{ color: "#000", backgroundColor: "white" }}
                    // iconProps={{ iconName: "Accept" }}
                    text="Delete"
                    onClick={() => {
                        let tempCon = [...this.props.tempContributors];
                        tempCon.splice(index, 1);
                        this.props.updateParentContributorState(tempCon);

                    }}
                />
            };
            default: {
                return <span>{fieldValue}</span>
            }
        }
    }
    public render(): ReactElement<IProps> {
        return (
            <div className={`ms-Grid-row`}>
                <div className={`ms-Grid-col ms-sm12`} style={{ paddingBottom: 10 }}>
                    <ShowADUserComponent isReadOnlyForm={this.props.isReadOnlyForm} sectionInfo={this.props.sectionInfo} allADUsers={this.props.allADUsers} updatePeopleCB={this.updateContributorState} sectionNumber={this.props.contributorPanelId} fieldState={"Contributor"} fieldName="" isMandatory={false}></ShowADUserComponent>
                </div>
                <div className={`ms-Grid-col ms-sm12`} style={{ paddingBottom: 10 }}>
                    <DetailsList
                        columns={contributorFormColumns}
                        items={this.props.tempContributors}
                        onRenderItemColumn={this.renderGridItems}
                        selectionMode={SelectionMode.none}
                        compact={true}
                    ></DetailsList>
                </div>
            </div>
        );
    }
}

export default ContributorDialog;

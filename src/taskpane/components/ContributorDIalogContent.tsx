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
import { contributorFormColumns, dialogContentProps, gridFormColumns } from "./SectionFormGridCont";

const MyComponent = () => {
    return <div></div>;
};

export interface IProps extends ComboboxProps {
    contributors: any[];
    updateParentContributorState: Function;
}

export interface IState {
}


export class ContributorDialog extends Component<IProps, IState> {
    constructor(props: IProps) {
        super(props);
        this.state = {
        };
    }
    public renderGridItems = (item: SectionAssignment, index: number, column: IColumn) => {
        const fieldValue = item[column.fieldName];
        switch (column.key) {
            case 'Action': {
                return <DefaultButton
                    style={{ color: "#000", backgroundColor: "white" }}
                    // iconProps={{ iconName: "Accept" }}
                    text="Delete"
                    onClick={() => {
                        let newContributors = this.props.contributors.map((item, i) => {
                            return index != i;
                        })
                        this.props.updateParentContributorState(newContributors);
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
                    {/* <ShowADUserComponent fieldState={"Secondary"} fieldName="" isMandatory={true}></ShowADUserComponent> */}
                </div>
                <div className={`ms-Grid-col ms-sm12`} style={{ paddingBottom: 10 }}>
                    <DetailsList
                        columns={contributorFormColumns}
                        items={this.props.contributors}
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

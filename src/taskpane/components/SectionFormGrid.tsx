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
import ContributorDialog from "./ContributorDIalogContent";

const MyComponent = () => {
    return <div></div>;
};

export interface IProps extends ComboboxProps {
    sections: SectionAssignment[];
    updateParentSectionState: Function;
}

export interface IState {
    contributorPanelId: number;
    tempContributors: any[];

}


export class SectionFormGrid extends Component<IProps, IState> {
    constructor(props: IProps) {
        super(props);
        this.state = {
            contributorPanelId: null,
            tempContributors: null
        };
    }

    public updateContributorState = (newContributors) => {
        this.setState({ tempContributors: newContributors });
    }
    public updateParentSectionState = () => {
        let tempSectionItem = { ...this.props.sections };
        // let tempSections = this.state.Sections.map(item => {
        //     if (item.SectionNumber == updatedSectionItem.SectionNumber)
        //         return item = updatedSectionItem;
        // });
        this.props.updateParentSectionState(tempSectionItem);
    }
    public updateContributors = () => {
        //update contributors in props
        this.setState({ contributorPanelId: null, tempContributors: null });
    }
    public renderGridItems = (item: SectionAssignment, index: number, column: IColumn) => {
        const fieldValue = item[column.fieldName];
        switch (column.key) {
            case 'SNo': {
                return <span>{item.SectionNumber}</span>
            };
            case 'Section': {
                return <span>{`Section ${item.SectionNumber}`}</span>
            };
            case 'Primary': {
                return <ShowADUserComponent fieldID={"Primary"} fieldName="" isMandatory={true}></ShowADUserComponent>
            };
            case 'Secondary': {
                return <ShowADUserComponent fieldID={"Secondary"} fieldName="" isMandatory={true}></ShowADUserComponent>
            };
            case 'TargetDate': {
                return <DatePicker
                    firstDayOfWeek={DayOfWeek.Sunday}
                    firstWeekOfYear={1}
                    showMonthPickerAsOverlay={true}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    value={item.DeadLineDate}
                    // onSelectDate={(date) => {
                    //     let tempSectionItem = { ...this.props.sectionInfo };
                    //     tempSectionItem.DeadLineDate = date;
                    // }}
                    // DatePicker uses English strings by default. For localized apps, you must override this prop.
                    strings={defaultDatePickerStrings}
                />
            };
            case 'Contributors': {
                return <span>{item.Contributor.join(", ")}</span>
            };
            case 'ManageContributors': {
                return <DefaultButton
                    style={{ color: "#000", backgroundColor: "white" }}
                    // iconProps={{ iconName: "Accept" }}
                    text="Manage"
                    onClick={() => {
                        this.setState({ contributorPanelId: index })
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
            <div>
                <DetailsList
                    columns={gridFormColumns}
                    items={this.props.sections}
                    onRenderItemColumn={this.renderGridItems}
                    selectionMode={SelectionMode.none}
                    compact={true}
                ></DetailsList>
                <Dialog
                    hidden={this.state.contributorPanelId == null}
                    onDismiss={() => { this.setState({ contributorPanelId: null, tempContributors: null }) }}
                    dialogContentProps={dialogContentProps}
                // modalProps={{isBlocking: false}}
                >
                    {this.state.contributorPanelId != null && <ContributorDialog contributors={this.props.sections.map((item, index) => { return index == this.state.contributorPanelId })} updateParentContributorState={this.updateContributorState}></ContributorDialog>}
                    <DialogFooter>
                        <PrimaryButton onClick={this.updateContributors} text="Ok" />
                        <DefaultButton onClick={() => {
                            this.setState({ contributorPanelId: null, tempContributors: null });
                        }} text="Cancel" />
                    </DialogFooter>
                </Dialog>
                {/* <h3>{`Section ${this.props.sectionInfo.SectionNumber}`}</h3>
                <div>
                    <ShowADUserComponent fieldName="Primary Owner" isMandatory={true}></ShowADUserComponent>
                </div>
                <div>
                    <ShowADUserComponent fieldName="Secondary Owner" isMandatory={true}></ShowADUserComponent>
                </div>
                <div>
                    <ShowADUserComponent fieldName="Contributor(s)" isMandatory={false}></ShowADUserComponent>
                </div>
                <div>
                    <DatePicker
                        label={"Target Date"}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        onSelectDate={(date) => {
                            let tempSectionItem = { ...this.props.sectionInfo };
                            tempSectionItem.DeadLineDate = date;
                        }}
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
                    />
                </div> */}
            </div>
        );
    }
}

export default SectionFormGrid;

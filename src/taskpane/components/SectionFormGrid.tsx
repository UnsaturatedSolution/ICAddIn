import React, { Component, ReactElement } from "react";
import type { ComboboxProps } from "@fluentui/react-components";
import { GetAllSiteUsersSSO, UpdateRequestSSO, getAllUsersSSO, getSearchUser } from "../../helpers/sso-helper";
import { DatePicker, DayOfWeek, DefaultButton, DetailsList, Dialog, DialogFooter, IColumn, Panel, Pivot, PivotItem, PrimaryButton, SelectionMode, TextField, Toggle, TooltipHost, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import ShowADUserComponent from "./CustomPicker";
import { contributorFormColumns, dialogContentProps, gridFormColumns, gridFormDoneColumn, gridFormManageColumn } from "../../constants/SectionFormGridCont";
import ContributorDialog from "./ContributorDIalogContent";
import { formatSectionName, mergeADSPUsers, onFormatDate } from "../../utilities/utility";
import * as appConst from "../../constants/appConst";

export interface IProps extends ComboboxProps {
    isReadOnlyForm: boolean;
    sections: SectionAssignment[];
    forceResetGrid: boolean;
    updateParentSectionState: Function;

}

export interface IState {
    contributorPanelId: number;
    tempContributors: any[];
    allADUsers: any[];

}

export class SectionFormGrid extends Component<IProps, IState> {
    constructor(props: IProps) {
        super(props);
        this.state = {
            contributorPanelId: null,
            tempContributors: [],
            allADUsers: []
        };
    }

    public componentDidMount(): void {
        getAllUsersSSO((result: any) => {
            const allADUsers = result.value;
            GetAllSiteUsersSSO((result: any) => {
                console.log(result);
                const allSPUsers = JSON.parse(result).d.results;
                const allUsers = mergeADSPUsers(allADUsers, allSPUsers);
                this.setState({ allADUsers: allUsers });
            });
        });
    }


    public updateParentSectionState = (selectedSectionNumber, propValObj) => {
        let tempSectionItems = this.props.sections.map(item => {
            let updatedItem = { ...item };
            if (item.SectionNumber == selectedSectionNumber)
                updatedItem = { ...updatedItem, ...propValObj };
            return updatedItem;
        });
        this.props.updateParentSectionState(tempSectionItems);
    }
    public saveContributors = () => {
        let tempSectionItems = [...this.props.sections]
        tempSectionItems[this.state.contributorPanelId].Contributors = this.state.tempContributors;
        this.setState({ contributorPanelId: null, tempContributors: [] }, () => {
            this.props.updateParentSectionState(tempSectionItems);
        });
    }

    public renderGridItems = (item: SectionAssignment, index: number, column: IColumn) => {
        const fieldValue = item[column.fieldName];
        switch (column.key) {
            case 'Done': {
                return <Toggle />
            };
            case 'SNo': {
                return <span>{item.SectionNumber}</span>
            };
            case 'Section': {
                return <TooltipHost
                    content={item.SectionName}
                    id={"SectionName"}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block' } }}
                >
                    <span style={{ whiteSpace: "break-spaces" }} aria-describedby={"SectionName"}>{`${formatSectionName(item.SectionName)}`}</span>
                </TooltipHost>
            };
            case 'Primary': {
                return <TooltipHost
                    content={item.POwnerDisplayName}
                    id={"PrimaryOwner"}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block', width: "100%" } }}
                >
                    {!this.props.isReadOnlyForm ?
                        <ShowADUserComponent isReadOnlyForm={this.props.isReadOnlyForm} forceResetGrid={this.props.forceResetGrid} sectionInfo={item} allADUsers={this.state.allADUsers} updatePeopleCB={this.updateParentSectionState} sectionNumber={item.SectionNumber} fieldState={"POwner"} fieldName="" isMandatory={true}></ShowADUserComponent>
                        : <span>{item.POwnerDisplayName}</span>}
                </TooltipHost>
            };
            case 'Secondary': {
                return <TooltipHost
                    content={item.SOwnerDisplayName}
                    id={"SecondaryOwner"}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block', width: "100%" } }}
                >
                    {!this.props.isReadOnlyForm ?
                        <ShowADUserComponent isReadOnlyForm={this.props.isReadOnlyForm} forceResetGrid={this.props.forceResetGrid} sectionInfo={item} allADUsers={this.state.allADUsers} updatePeopleCB={this.updateParentSectionState} sectionNumber={item.SectionNumber} fieldState={"SOwner"} fieldName="" isMandatory={true}></ShowADUserComponent>
                        : <span>{item.SOwnerDisplayName}</span>}
                </TooltipHost>
            };
            case 'TargetDate': {
                return <TooltipHost
                    content={onFormatDate(item.DeadLineDate)}
                    id={"DeadLineDate"}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block' } }}
                >
                    <DatePicker
                        disabled={this.props.isReadOnlyForm}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        formatDate={onFormatDate}
                        value={item.DeadLineDate}
                        onSelectDate={(date) => {
                            this.updateParentSectionState(item.SectionNumber, { ...item, ...{ "DeadLineDate": date } })
                        }}
                        strings={defaultDatePickerStrings}
                    />
                </TooltipHost>
            };
            case 'Contributors': {
                return <TooltipHost
                    content={item.Contributors.map(contributor => contributor.ContributorDisplayName).join(", ")}
                    id={"Contributors"}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block' } }}
                >
                    <span style={{ whiteSpace: "break-spaces" }}>{item.Contributors.map(contributor => contributor.ContributorDisplayName).join(", ")}</span>
                </TooltipHost>

            };
            case 'ManageContributors': {
                return !this.props.isReadOnlyForm ? <DefaultButton
                    style={{ color: "#000", backgroundColor: "white" }}
                    text="Manage"
                    onClick={() => {
                        this.setState({ contributorPanelId: index, tempContributors: item.Contributors });
                    }}
                /> : null;
            };
            default: {
                return <TooltipHost
                    content={fieldValue}
                    id={column.key}
                    calloutProps={{ gapSpace: 0 }}
                    styles={{ root: { display: 'inline-block' } }}
                >
                    <span>{fieldValue}</span>
                </TooltipHost>

            }
        }
    }
    public render(): ReactElement<IProps> {
        const contributorPanelSectionInfo: SectionAssignment[] = this.props.sections.filter((item, index) => { return index == this.state.contributorPanelId })
        return (
            <div className={`ms-Grid-col ms-sm12`}>
                <DetailsList
                    columns={this.props.isReadOnlyForm ? [...gridFormDoneColumn, ...gridFormColumns] : [...gridFormColumns, ...gridFormManageColumn]}
                    items={this.props.sections}
                    onRenderItemColumn={this.renderGridItems}
                    selectionMode={SelectionMode.none}
                    compact={true}
                ></DetailsList>
                <Panel
                    isOpen={this.state.contributorPanelId != null}
                    onDismiss={() => { this.setState({ contributorPanelId: null, tempContributors: [] }) }}
                    headerText="Manage Contributors"
                    closeButtonAriaLabel="Close"
                    onRenderFooterContent={() => {
                        return <div>
                            <PrimaryButton style={{ paddingRight: "10px" }} onClick={this.saveContributors} text="OK" />
                            <DefaultButton onClick={() => {
                                this.setState({ contributorPanelId: null, tempContributors: [] });
                            }} text="Cancel" />
                        </div>
                    }}
                    isFooterAtBottom={true}
                >
                    {this.state.contributorPanelId != null &&
                        <ContributorDialog
                            contributorPanelId={this.state.contributorPanelId}
                            sectionInfo={contributorPanelSectionInfo ? contributorPanelSectionInfo[0] : null}
                            updateParentContributorState={(tempContributors) => {
                                this.setState({ tempContributors: tempContributors })
                            }}
                            isReadOnlyForm={this.props.isReadOnlyForm}
                            tempContributors={this.state.tempContributors}
                            allADUsers={this.state.allADUsers}
                        ></ContributorDialog>
                    }
                </Panel>
            </div >
        );
    }
}

export default SectionFormGrid;

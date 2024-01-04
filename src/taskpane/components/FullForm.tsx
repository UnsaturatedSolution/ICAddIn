import React, { Component, ReactElement } from "react";
import { Textarea, type ComboboxProps } from "@fluentui/react-components";
import { DatePicker, DayOfWeek, DefaultButton, TextField, Toggle, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import SectionFormComponent from "./SectionForm";
import { CreateRequestSSO, GetSPDataSSO, GetSPDocSSO, UpdateRequestSSO } from "../../helpers/sso-helper";
import { FlagPrideIntersexInclusiveProgress20Filled } from "@fluentui/react-icons";
import SectionFormGrid from "./SectionFormGrid";
import { onFormatDate } from "../../utilities/utility";
import * as appConst from "../../constants/appConst";

const MyComponent = () => {
    return <div></div>;
};

export interface IProps extends ComboboxProps {
    officeContext: any;
    docGUID: string;
    sectionInfo: any[];
    docInfo: any;
    updateDocInfo: Function;
}

export interface IState {
    options: any[];
    Sections: SectionAssignment[];
    isDocFreezed: boolean;
    searchText: string;
    tempItemID: string;
    docDueDate: Date;
    forceResetGrid: boolean;
}

let sampleCreateItem = {
    Title: "API Created Item",
    SectionName: "Section 1",
    SectionSequence: 1,
    // PrimaryOwner: "Harsha@vc.com",
    // SecondaryOwner: "Harsha@vc.com",
    // Contributors: "Harsha@vc.com",
    Status: "NotStarted",
    DocumentID: "0",
    Comments: "Comms",
    TargetDate: new Date(),
    SectionID: "ABC123"
}
export class FullFormComponent extends Component<IProps, IState> {
    constructor(props: IProps) {
        super(props);
        this.state = {
            options: [],
            Sections: [],
            isDocFreezed: false,
            searchText: "",
            tempItemID: "0",
            docDueDate: null,
            forceResetGrid: false
        };
    }
    async componentDidMount(): Promise<void> {
        this.setState({ Sections: this.props.sectionInfo });
    }
    async componentDidUpdate(prevProps, prevState): Promise<void> {
        const prevSectionRefs = this.getSectionRefStr(prevProps.sectionInfo);
        const currectionRefs = this.getSectionRefStr(this.props.sectionInfo);
        if (prevSectionRefs != currectionRefs) {
            this.setState({ Sections: this.props.sectionInfo });
        }
    }
    public getSectionRefStr = (sectionInfo) => {
        let refStr = "";
        if (sectionInfo.length > 0) {
            refStr = sectionInfo.map(item => item.SectionID).join(";");
        }
        return refStr;
    }
    public mapSectionRowToItem = (sectionInfo) => {
        return sectionInfo.map((item, index) => {
            const constributorsId = item.Contributors.map(contributor => contributor.ContributorID);
            return {
                SectionSequence: item.SectionNumber,
                SectionName: item.SectionName ? item.SectionName : "",
                PrimaryOwnerId: item.POwnerID ? item.POwnerID : 0,
                SecondaryOwnerId: item.SOwnerID ? item.SOwnerID : 0,
                Status: "InProgress",
                TargetDate: item.DeadLineDate ? new Date(item.DeadLineDate) : null,
                DocumentID: item.DocumentID ? item.DocumentID : this.props.docGUID,
                SectionID: item.SectionID ? item.SectionID : "",
                ContributorsId: constributorsId
            }
        })
    }
    public syncSection = async () => {
        const tempSections = [...this.state.Sections];
        await Word.run(async (context) => {
            const secs = context.document.sections;
            secs.load('body/style');
            await context.sync();
            const sectionRefs = this.getSectionRefStr(tempSections).split(";");
            let Items = [];

            for (let i = 0; i < secs.items.length; i++) {
                let section = secs.items[i];
                let index = i;
                const sectionRefID = section["__R"] ? section["__R"] : "";
                const itemRefIDIndex = sectionRefs.indexOf(sectionRefID);
                const bodyObj = secs.items[index].body;
                bodyObj.load('text');
                await context.sync();
                console.log(bodyObj);
                const sectionNameArr = bodyObj.text.trim().split("\r");
                const sectionName = sectionNameArr.length > 0 ? sectionNameArr[0] : "";
                if (itemRefIDIndex >= 0) {
                    Items.push({ ...tempSections[itemRefIDIndex], ...{ SectionNumber: index + 1 } });
                }
                else {
                    Items.push({
                        SectionNumber: index + 1,
                        SectionName: sectionName,
                        POwnerID: 0,
                        POwnerEmail: "",
                        POwnerDisplayName: "",
                        SOwnerID: 0,
                        SOwnerEmail: "",
                        SOwnerDisplayName: "",
                        Contributors: [],
                        DeadLineDate: "",
                        DocumentID: this.props.docGUID,
                        SectionID: sectionRefID
                    });
                }
            }
            // );
            this.setState({ Sections: Items, forceResetGrid: true }, () => {
                this.setState({ forceResetGrid: false });
            });
        });
    };
    public resetAllSections = () => {
        let tempSections = this.state.Sections.map((section) => {
            return {
                ...section,
                POwnerID: 0,
                POwnerEmail: "",
                POwnerDisplayName: "",
                SOwnerID: 0,
                SOwnerEmail: "",
                SOwnerDisplayName: "",
                Contributor: [],
                DeadLineDate: null,
            };
        })
        this.setState({ Sections: tempSections, forceResetGrid: true }, () => {
            this.setState({ forceResetGrid: false });
        });

    }
    private addNewSection=()=>{
        let tempSectipons = [...this.state.Sections];
        tempSectipons.push({
            itemID:0,
            SectionNumber: this.state.Sections.length+1,
            SectionName: "",
            POwnerID: 0,
            POwnerEmail: "",
            POwnerDisplayName: "",
            SOwnerID: 0,
            SOwnerEmail: "",
            SOwnerDisplayName: "",
            Contributors: [],
            DeadLineDate: null,
            DocumentID: this.props.docGUID,
            SectionID: "NA"
        });
        this.setState({Sections:tempSectipons})
    }
    public updateSectionState = (sections) => {
        this.setState({ Sections: sections });
    }
    public createAllSections = async () => {
        let docDetailsItem = {
            DocumentID: this.props.docGUID,
            DocStatus: "Initiated",
            DueDate: this.state.docDueDate,
            ShouldFreezeDoc: this.state.isDocFreezed
        };

        let docResponse = await CreateRequestSSO(appConst.lists.documentDetails, docDetailsItem);
        if (docResponse) {
            this.props.updateDocInfo({
                DocStatus: "Initiated",
                DueDate: this.state.docDueDate,
                ShouldFreezeDoc: true
            })
            const mappedSections = this.mapSectionRowToItem(this.state.Sections);
            await Promise.all(
                mappedSections.map(async (section, index) => {
                    let response = await CreateRequestSSO(appConst.lists.assigneeDetails, section);
                    if (!response) {
                        throw new Error("Middle tier didn't respond");
                    } else if (response.claims) {
                        console.log("data saved");
                    }
                })
            ).then(values => {
                console.log(values);
            });
        }
    }

    // public createSectionItem = async () => {
    //     let createItem = { ...sampleCreateItem };
    //     createItem.DocumentID = this.props.docGUID;
    //     let response: any = await CreateRequestSSO(createItem);
    //     if (!response) {
    //         throw new Error("Middle tier didn't respond");
    //     } else if (response.claims) {
    //         console.log("data saved");
    //     }
    // }
    // public updateSectionItem = async () => {
    //     let itemID = `${this.state.tempItemID}`;
    //     const constributorsId = 
    //     let updateItem = {ContributorsId: { "results": ["36","42"] }
    //     }
    //     let response: any = await UpdateRequestSSO(createItem, itemID,appConst.lists.assigneeDetails);
    //     if (!response) {
    //         throw new Error("Middle tier didn't respond");
    //     } else if (response.claims) {
    //         console.log("data saved");
    //     }
    // }
    // public updateContributors = async () => {
    //     const contributorPanelSectionInfo: SectionAssignment[] = this.props.sections.filter((item, index) => { return index == this.state.contributorPanelId });
    //     const constributorsId = contributorPanelSectionInfo[0].Contributors.map(contributor=>`${contributor.ContributorID}`);
    //     let response: any = await UpdateRequestSSO({ContributorsId:{ "results": constributorsId }}, contributorPanelSectionInfo[0].itemID,appConst.lists.assigneeDetails);
    //     if (!response) {
    //         throw new Error("Middle tier didn't respond");
    //     } else if (response.claims) {
    //         console.log("data saved");
    //     }
    // }
    public deleteSectionItem = async () => {
        let itemID = `${this.state.tempItemID}`;
        let createItem = {
            IsActive: false
        }
        let response: any = await UpdateRequestSSO(createItem, itemID, appConst.lists.assigneeDetails);
        if (!response) {
            throw new Error("Middle tier didn't respond");
        } else if (response.claims) {
            console.log("data saved");
        }
    }
    public render(): ReactElement<IProps> {
        const isDocInitiated = this.props.docInfo.DocStatus == "Initiated";
        // const isDocInitiated =false;
        return (
            <div className={`ms-Grid-row`} style={{ padding: 20 }}>
                <p>{`Document Status : ${this.props.docInfo.DocStatus ? this.props.docInfo.DocStatus : "Not Started"}`}</p>
                {(!isDocInitiated && this.state.Sections.length>0) && <div className={`ms-Grid-col ms-sm12`} style={{ marginTop: 10, display: "flex", alignItems: 'center', justifyContent: 'space-between' }}>
                    <DefaultButton
                        style={{ color: "#000", backgroundColor: "white" }}
                        disabled={!(this.state.Sections && this.state.Sections.length > 0)}
                        text="Initiate Document"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.createAllSections}
                    />
                    <DefaultButton
                        style={{ color: "#000", backgroundColor: "white" }}
                        disabled={!(this.state.Sections && this.state.Sections.length>0)}
                        text="Add New Section"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.addNewSection}
                    />
                    <DefaultButton
                        style={{ color: "#000", backgroundColor: "white" }}
                        disabled={!(this.state.Sections && this.state.Sections.length > 0)}
                        text="Save"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.createAllSections}
                    />
                    <DefaultButton
                        style={{ color: "#000", backgroundColor: "white" }}
                        disabled={!(this.state.Sections && this.state.Sections.length > 0)}
                        text="Reset"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.resetAllSections}
                    />
                </div>}
                <div className={`ms-Grid-col ms-sm12`} style={{ display: "flex", alignItems: 'center', justifyContent: 'space-between' }}>
                    <DatePicker
                        label={"Due Date"}
                        disabled={isDocInitiated}
                        // isRequired={true}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        formatDate={onFormatDate}
                        strings={defaultDatePickerStrings}
                        value={this.props.docInfo.DueDate ? new Date(this.props.docInfo.DueDate) : this.state.docDueDate}
                        onSelectDate={(date) => {
                            let tempSections = this.state.Sections.map(sectionItem => {
                                if (sectionItem.DeadLineDate == null || sectionItem.DeadLineDate.toString() == "") {
                                    return { ...sectionItem, ...{ DeadLineDate: date } };
                                }
                                else
                                    return sectionItem;
                            });
                            this.setState({ docDueDate: date, Sections: tempSections });
                        }}
                    />
                    <Toggle label="Freeze" checked={this.state.isDocFreezed} onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => { this.setState({ isDocFreezed: checked }) }} />
                    {!isDocInitiated ? <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        text="Sync Section"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.syncSection}
                    /> : null}
                </div>
                
                <SectionFormGrid
                    isReadOnlyForm={isDocInitiated}
                    sections={this.state.Sections}
                    forceResetGrid={this.state.forceResetGrid}
                    updateParentSectionState={this.updateSectionState}
                ></SectionFormGrid>
                
            </div>
        );
    }
}

export default FullFormComponent;

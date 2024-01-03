import React, { Component, ReactElement } from "react";
import { Textarea, type ComboboxProps } from "@fluentui/react-components";
import { DatePicker, DayOfWeek, DefaultButton, TextField, Toggle, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import SectionFormComponent from "./SectionForm";
import { CreateRequestSSO, GetSPDataSSO, GetSPDocSSO, UpdateRequestSSO } from "../../helpers/sso-helper";
import { FlagPrideIntersexInclusiveProgress20Filled } from "@fluentui/react-icons";
import SectionFormGrid from "./SectionFormGrid";
import { onFormatDate } from "../../utilities/utility";


const MyComponent = () => {
    return <div></div>;
};

export interface IProps extends ComboboxProps {
    officeContext: any;
    docGUID: string;
    sectionInfo: any[];
    docInfo: any;
}

export interface IState {
    options: any[];
    Sections: SectionAssignment[];
    isDocFreezed: boolean;
    searchText: string;
    // docGUID: string;
    tempItemID: string;
    docDueDate: Date;
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
            // docGUID: "",
            tempItemID: "0",
            docDueDate: null
        };
    }
    async componentDidMount(): Promise<void> {
        // await this.getDocumentMetadata();
        // await this.getSPData();
        // await this.syncSection();
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
    // public mapSectionItemToRow = (sectionInfo) => {
    //     return sectionInfo.map((item, index) => {
    //         return {
    //             SectionNumber: item.SectionSequence,
    //             SectionName: item.SectionName ? item.SectionName : "",
    //             POwnerID: item.PrimaryOwnerId ? item.PrimaryOwnerId : 0,
    //             POwnerEmail: "",
    //             SOwnerID: item.SecondaryOwnerId ? item.SecondaryOwnerId : 0,
    //             SOwnerEmail: "",
    //             Contributors: [],
    //             DeadLineDate: item.TargetDate ? new Date(item.TargetDate) : null,
    //             DocumentID: item.DocumentID ? item.DocumentID : this.props.docGUID,
    //             SectionID: item.SectionID ? item.SectionID : ""
    //         }
    //     })
    // }
    public mapSectionRowToItem = (sectionInfo) => {
        return sectionInfo.map((item, index) => {
            return {
                SectionSequence: item.SectionNumber,
                SectionName: item.SectionName ? item.SectionName : "",
                PrimaryOwnerId: item.POwnerID ? item.POwnerID : 0,
                // POwnerEmail: "",
                // SOwnerID: 0,
                SecondaryOwnerId: item.SOwnerID ? item.SOwnerID : 0,
                // SOwnerEmail: "",
                // Contributors: [],
                Status: "NotStarted",
                // Comments: "Comms",
                TargetDate: item.DeadLineDate ? new Date(item.DeadLineDate) : null,
                DocumentID: item.DocumentID ? item.DocumentID : this.props.docGUID,
                SectionID: item.SectionID ? item.SectionID : ""
            }
        })
    }
    // public getData = async () => {
    //     await Word.run(async (context) => {
    //         const mySections = context.document.sections;
    //         mySections.load('body/style');
    //         await context.sync();
    //         console.log(mySections);
    //         const firstbody = mySections.items[0].body;
    //         firstbody.load('text');
    //         await context.sync();
    //         const sectionName = `${firstbody.text.split(" ")[0]} ${firstbody.text.split(" ")[1]} ${firstbody.text.split(" ")[2]}`
    //         console.log(sectionName);
    //         console.log("Added a header to the first section.");
    //         this.syncSection();
    //     });
    // }
    public syncSection = async () => {
        const tempSections = [...this.state.Sections];
        await Word.run(async (context) => {
            const secs = context.document.sections;
            secs.load('body/style');
            // const docProps = context.document.properties;
            // context.load(docProps);
            // context.load(docProps.customProperties);
            // context.load(secs);
            await context.sync();
            const sectionRefs = this.getSectionRefStr(tempSections).split(";");
            let Items = [];

            // await Promise.all(
            // secs.items.map(async (section, index) => {
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
                // const sectionItem = secs.items[index].body;
                // sectionItem.load('text');
                // await context.sync();
                // console.log(sectionItem);
                // sectionItem.load('text');
                // await context.sync();
                // console.log(sectionItem);
                // let body = secs.items[index].body;
                // context.load(body);
                // await context.sync();
                // body.load("text");
                // await context.sync();
                if (itemRefIDIndex >= 0) {
                    Items.push({ ...tempSections[itemRefIDIndex], ...{ SectionNumber: index + 1 } });
                }
                else {
                    Items.push({
                        SectionNumber: index + 1,
                        SectionName: sectionName,
                        POwnerID: 0,
                        POwnerEmail: "",
                        POwneDisplayName: "",
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
            this.setState({ Sections: Items });
        });
    };
    public resetAllSections = () => {
        let tempSections = this.state.Sections.map((section) => {
            return {
                ...section,
                POwnerID: 0,
                POwnerEmail: "",
                POwneDisplayName: "",
                SOwnerID: 0,
                SOwnerEmail: "",
                SOwnerDisplayName: "",
                Contributor: [],
                DeadLineDate: null,
            };
        })
        this.setState({ Sections: tempSections });

    }
    // private addNewSection=()=>{
    //     let tempSectipons = [...this.state.Sections];
    //     tempSectipons.push({
    //         SectionNumber: this.state.Sections.length+1,
    //         SectionName: "",
    //         POwnerID: 0,
    //         POwnerEmail: "",
    //         POwneDisplayName: "",
    //         SOwnerID: 0,
    //         SOwnerEmail: "",
    //         SOwnerDisplayName: "",
    //         Contributor: [],
    //         DeadLineDate: null,
    //         DocumentID: this.props.docGUID,
    //         SectionID: "NA"
    //     });
    //     this.setState({Sections:tempSectipons})
    // }
    public updateSectionState = (sections) => {
        // let tempSections = this.state.Sections.map(item => {
        //     if (item.SectionNumber == updatedSectionItem.SectionNumber)
        //         return item = updatedSectionItem;
        // });
        this.setState({ Sections: sections });
    }
    // public getDocumentMetadata = async () => {
    //     let fileURL = this.props.officeContext.document.url;
    //     if (fileURL.indexOf('sharepoint.com') > -1) {
    //         // let docName = fileURL.split('/')[fileURL.split('/').length - 1];
    //         let serverRelativeUrl = fileURL.split('https://vichitra.sharepoint.com')[fileURL.split('https://vichitra.sharepoint.com').length - 1];

    //         let response: any = await GetSPDocSSO(serverRelativeUrl, {});
    //         console.log(response);
    //         if (!response) {
    //             throw new Error("Middle tier didn't respond");
    //         } else if (response.claims) {
    //             this.setState({ docGUID: response.UniqueId });
    //         }
    //     }
    // }
    // public getSPData = async () => {
    //     let docID = "";
    //     let response: any = await GetSPDataSSO(`Title ne ""`, {});
    //     console.log(response);
    //     if (!response) {
    //         throw new Error("Middle tier didn't respond");
    //     } else if (response.claims) {
    //         console.log("data saved");
    //     }
    // }
    public createAllSections = async () => {
        let docDetailsItem = {
            DocumentID: this.props.docGUID,
            DocStatus: "Initiated",
            DueDate: this.state.docDueDate,
            ShouldFreezeDoc: true
        };
        CreateRequestSSO("InvestCorpDocumentDetails", docDetailsItem);
        const mappedSections = this.mapSectionRowToItem(this.state.Sections);
        await Promise.all(
            mappedSections.map(async (section, index) => {
                let response: any = await CreateRequestSSO("InvestcorpDocumentAssignees", section);
                if (!response) {
                    throw new Error("Middle tier didn't respond");
                } else if (response.claims) {
                    console.log("data saved");
                }
            })
        ).then(values => {
            console.log(values);
        });
        // for (let i = 0; i < mappedSections.length; i++) {
        //     const sectionItem = mappedSections[i];
        //     let response: any = await CreateRequestSSO(sectionItem);
        //     if (!response) {
        //         throw new Error("Middle tier didn't respond");
        //     } else if (response.claims) {
        //         console.log("data saved");
        //     }
        // }
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
    public updateSectionItem = async () => {
        let itemID = `${this.state.tempItemID}`;
        let createItem = {
            Title: "API Created Item Updated",
            SectionName: "Section 1",
            SectionSequence: 1,
            // PrimaryOwner: "Harsha@vc.com",
            // SecondaryOwner: "Harsha@vc.com",
            // Contributors: "Harsha@vc.com",
            Status: "NotStarted",
            // DocumentID: 12,
            Comments: "Comms",
            TargetDate: new Date(),
            SectionID: "ABC123",
            DocumentID: this.props.docGUID
        }
        let response: any = await UpdateRequestSSO(createItem, itemID, 'InvestcorpDocumentAssignees');
        if (!response) {
            throw new Error("Middle tier didn't respond");
        } else if (response.claims) {
            console.log("data saved");
        }
    }
    public deleteSectionItem = async () => {
        let itemID = `${this.state.tempItemID}`;
        let createItem = {
            IsActive: FlagPrideIntersexInclusiveProgress20Filled
        }
        let response: any = await UpdateRequestSSO(createItem, itemID,'InvestcorpDocumentAssignees');
        if (!response) {
            throw new Error("Middle tier didn't respond");
        } else if (response.claims) {
            console.log("data saved");
        }
    }
    public render(): ReactElement<IProps> {
        const isDocInitiated = this.props.docInfo.DocStatus == "Initiated";
        return (
            <div className={`ms-Grid-row`} style={{ padding: 20 }}>
                <p>{`Document Status : ${this.props.docInfo.DocStatus ? this.props.docInfo.DocStatus : ""}`}</p>
                <div className={`ms-Grid-col ms-sm12`} style={{ display: "flex", alignItems: 'center', justifyContent: 'space-between' }}>
                    <DatePicker
                        label={"Due Date"}
                        isRequired={true}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        formatDate={onFormatDate}
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
                        value={this.props.docInfo.DueDate?new Date(this.props.docInfo.DueDate) : this.state.docDueDate}
                        onSelectDate={(date) => {
                            this.setState({ docDueDate: date });
                        }}
                    />
                    <Toggle label="Freeze" checked={this.state.isDocFreezed} onChange={(event: React.MouseEvent<HTMLElement>, checked?: boolean) => { this.setState({ isDocFreezed: checked }) }} />
                    <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        text="Sync Section"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.syncSection}
                    />
                    {/* <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        iconProps={{ iconName: "Accept" }}
                        text="Test Item Create"
                        onClick={this.createSectionItem}
                    />
                    <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        iconProps={{ iconName: "Accept" }}
                        text="Test Update last item"
                        onClick={this.updateSectionItem}
                    />
                    <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        iconProps={{ iconName: "Accept" }}
                        text="Test delete last item"
                        onClick={this.deleteSectionItem}
                    /> */}
                </div>
                {/* <div className={`ms-Grid-col ms-sm12`} style={{ paddingBottom: 10 }}>
                    <TextField label="Filter" placeholder="Please provide text" value={this.state.searchText} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setState({ searchText: newValue })} iconProps={{ iconName: 'Search' }} />
                </div> */}
                {/* <div className={`ms-Grid-col ms-sm12`}>
                    {this.state.Sections.length > 0 && this.state.Sections.map((sectionItem) => {
                        return <SectionFormComponent
                            sectionInfo={sectionItem}
                            updateParentSectionState={this.updateSectionState}
                        ></SectionFormComponent>
                    })}
                </div> */}
                <SectionFormGrid
                    isReadOnlyForm={isDocInitiated}
                    sections={this.state.Sections}
                    updateParentSectionState={this.updateSectionState}
                ></SectionFormGrid>
                <div className={`ms-Grid-col ms-sm12`} style={{ marginTop: 10, display: "flex", alignItems: 'center', justifyContent: 'space-between' }}>
                    <DefaultButton
                        style={{ color: "#000", backgroundColor: "white" }}
                        disabled={!(this.state.Sections && this.state.Sections.length > 0)}
                        text="Initiate Document"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.createAllSections}
                    />
                    {/* <DefaultButton
                        style={{ color: "#000", backgroundColor: "white" }}
                        disabled={!(this.state.Sections && this.state.Sections.length>0)}
                        text="Add New Section"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.addNewSection}
                    /> */}
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
                </div>
            </div>
        );
    }
}

export default FullFormComponent;

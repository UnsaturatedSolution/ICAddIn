import React, { Component, ReactElement } from "react";
import { Textarea, type ComboboxProps } from "@fluentui/react-components";
import { DatePicker, DayOfWeek, DefaultButton, TextField, Toggle, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import SectionFormComponent from "./SectionForm";
import { CreateRequestSSO, DeleteRequestSSO, GetSPDataSSO, GetSPDocSSO, UpdateRequestSSO } from "../../helpers/sso-helper";
import { FlagPrideIntersexInclusiveProgress20Filled } from "@fluentui/react-icons";
import SectionFormGrid from "./SectionFormGrid";


const MyComponent = () => {
    return <div></div>;
};

export interface IProps extends ComboboxProps {
    officeContext: any;
    docGUID: string;
}

export interface IState {
    options: any[];
    Sections: SectionAssignment[];
    isDocFreezed: boolean;
    searchText:string;
    // docGUID: string;
    tempItemID: string;
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
            searchText:"",
            // docGUID: "",
            tempItemID: "0"
        };
    }
    async componentDidMount(): Promise<void> {
        // await this.getDocumentMetadata();
        // await this.getSPData();
        await this.syncSection();
    }
    public syncSection = async () => {
        console.log(this.props.officeContext);
        await Word.run(async (context) => {
            const secs = context.document.sections;
            const docProps = context.document.properties;
            context.load(docProps);
            context.load(docProps.customProperties);
            context.load(secs);
            await context.sync();
            console.log(secs);
            console.log(docProps);
            console.log(secs.toJSON().items);
            let Items = [];
            let sectionData = [];
            let tempData = null;
            await Promise.all(
                secs.toJSON().items.map((section, index) => {
                    Items.push({
                        SectionNumber: index + 1,
                        POwnerID: 0,
                        POwnerEmail: "",
                        SOwnerID: 0,
                        SOwnerEmail: "",
                        Contributor: [],
                        DeadLineDate: new Date(),
                        DocumentID: this.props.docGUID
                    });
                    sectionData.push({ section });
                })
            ).then(values => {
                tempData = values;
            });
            console.log(sectionData);
            this.setState({ Sections: Items });
        });
    };
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
        await Promise.all(
            this.state.Sections.map(async (section, index) => {
                let response: any = await CreateRequestSSO(section);
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

    public createSectionItem = async () => {
        let createItem = { ...sampleCreateItem };
        createItem.DocumentID = this.props.docGUID;
        let response: any = await CreateRequestSSO(createItem);
        if (!response) {
            throw new Error("Middle tier didn't respond");
        } else if (response.claims) {
            console.log("data saved");
        }
    }
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
        let response: any = await UpdateRequestSSO(createItem, itemID);
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
        let response: any = await UpdateRequestSSO(createItem, itemID);
        if (!response) {
            throw new Error("Middle tier didn't respond");
        } else if (response.claims) {
            console.log("data saved");
        }
    }
    public render(): ReactElement<IProps> {
        return (
            <div className={`ms-Grid-row`} style={{ padding: 20 }}>
                {/* <div className={`ms-Grid-col ms-sm12`} style={{ display: "flex" }}>
                    <Textarea size="large" value={this.state.tempItemID} onChange={(event) => {
                        this.setState({
                            tempItemID: event.target.value
                        })
                    }} />

                </div> */}
                <div className={`ms-Grid-col ms-sm12`} style={{ display: "flex",alignItems: 'center',justifyContent: 'space-between' }}>
                    <DatePicker
                        label={"Due Date"}
                        isRequired={true}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
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
                <div className={`ms-Grid-col ms-sm12`} style={{ paddingBottom:10}}>
                    <TextField label="Filter" placeholder="Please provide text" value={this.state.searchText} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string)=>this.setState({searchText:newValue})} iconProps={{ iconName: 'Search' }} />
                </div>
                {/* <div className={`ms-Grid-col ms-sm12`}>
                    {this.state.Sections.length > 0 && this.state.Sections.map((sectionItem) => {
                        return <SectionFormComponent
                            sectionInfo={sectionItem}
                            updateParentSectionState={this.updateSectionState}
                        ></SectionFormComponent>
                    })}
                </div> */}
                <SectionFormGrid sections={this.state.Sections}
                    updateParentSectionState={this.updateSectionState} ></SectionFormGrid>
                <div className={`ms-Grid-col ms-sm12`}>
                    <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        iconProps={{ iconName: "Accept" }}
                        text="Freeze Document"
                        onClick={() => { }}
                    />
                    <DefaultButton
                        style={{ marginLeft: 10, color: "#000", backgroundColor: "white" }}
                        text="Create All Sections"
                        iconProps={{ iconName: "Reply" }}
                        onClick={this.createAllSections}
                    />
                </div>
            </div>
        );
    }
}

export default FullFormComponent;

import React, { Component, ReactElement } from "react";
import {
  Combobox,
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
import { GetSPDocSSO, getSearchUser, GetSPListData, checkGroup, currentuserDetails } from "../../helpers/sso-helper";
import { DatePicker, DayOfWeek, DefaultButton, Pivot, PivotItem, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import SectionFormComponent from "./SectionForm";
import FullFormComponent from "./FullForm";
import SectionDetails from "./SectionDetails";
import * as appConst from "../../constants/appConst"

const MyComponent = () => {
  return <div></div>;
};

export interface IProps extends ComboboxProps {
  officeContext: any;
}

export interface IState {
  docGUID: string;
  // options: any[];
  // Sections: SectionAssignment[];
  SectionsDetails: any[];
  mappedSectionInfo: any[];
  docInfo: any;
  currentUserEmail: string;
}

export class HomeScreenComponent extends Component<IProps, IState> {
  constructor(props: IProps) {
    super(props);
    this.state = {
      docGUID: "",
      SectionsDetails: [],
      mappedSectionInfo: [],
      docInfo: {},
      currentUserEmail: ""
      // options: [],
      // Sections: [],
    };
  }
  // componentDidMount(): void {
  //   this.syncSection();
  // }
  // public syncSection = async () => {
  //   await Word.run(async (context) => {
  //     const secs = context.document.sections;
  //     const docProps = context.document.properties;
  //     context.load(docProps);
  //     context.load(docProps.customProperties);
  //     context.load(secs);
  //     await context.sync();
  //     console.log(secs.toJSON().items);
  //     let Items = [];
  //     await Promise.all(
  //       secs.toJSON().items.map(async (section, index) => {
  //         Items.push({
  //           SectionNumber: index + 1,
  //           POwnerID: 0,
  //           POwnerEmail: "",
  //           SOwnerID: 0,
  //           SOwnerEmail: "",
  //           Contributor: [],
  //           DeadLineDate: new Date(),
  //         });
  //       })
  //     );
  //     this.setState({ Sections: Items });
  //   });
  // };
  async componentDidMount(): Promise<void> {
    let tempState = { ...this.state };
    const currEmail = await this.currentUserDetails();
    tempState = { ...tempState, ...{ currentUserEmail: currEmail } };
    await this.checkUser();
    const docGUID = await this.getDocumentMetadata();
    tempState = { ...tempState, ...{ docGUID: docGUID } };
    const docInfo = await this.GetSPDocDetails(docGUID);
    if (docInfo) {
      tempState = { ...tempState, ...{ docInfo: docInfo } };
      const sectionInfo = await this.GetSPAssigneeData(docGUID);
      const mappedSectionInfo = this.mapSectionItemToRow(sectionInfo, docGUID)
      const sortedSectionInfo = mappedSectionInfo.sort((a, b) => a.SectionNumber - b.SectionNumber)
      tempState = { ...tempState, ...{ SectionsDetails: sectionInfo, mappedSectionInfo: sortedSectionInfo } };
    }
    this.setState(tempState);
  }
  public mapSectionItemToRow = (sectionInfo, docGUID) => {
    return sectionInfo.map((item, index) => {
      return {
        SectionNumber: item.SectionSequence,
        SectionName: item.SectionName,
        POwnerID: item.PrimaryOwnerId ? item.PrimaryOwnerId : 0,
        POwnerDisplayName: item.PrimaryOwner.Title,
        POwnerEmail: "",
        SOwnerID: item.SecondaryOwnerId ? item.SecondaryOwnerId : 0,
        SOwnerDisplayName: item.SecondaryOwner.Title,
        SOwnerEmail: "",
        Contributors: [],
        DeadLineDate: item.TargetDate ? new Date(item.TargetDate) : "",
        DocumentID: item.DocumentID ? item.DocumentID : docGUID,
        SectionID: item.SectionID ? item.SectionID : ""
      }
    })
  }
  public GetSPDocDetails = async (docGUID) => {
    const filter = `DocumentID eq '${docGUID}'`;
    let response = await GetSPListData("InvestCorpDocumentDetails", "*", "", filter);
    const result = JSON.parse(response);
    let docDetails = {};
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else {
      docDetails = response && result.d.results.length > 0 ? result.d.results[0] : {};
    }
    return docDetails;
  }
  public GetSPAssigneeData = async (docGUID) => {
    const filter = `DocumentID eq '` + docGUID + `' and IsActive eq 1`;
    let response: any = await GetSPListData("InvestcorpDocumentAssignees", "*,PrimaryOwner/Title,SecondaryOwner/Title", "PrimaryOwner,SecondaryOwner", filter);
    const result = JSON.parse(response);
    let sectionInfo = [];
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else {
      sectionInfo = response ? result.d.results : [];
    }
    return sectionInfo;
  }
  public getDocumentMetadata = async () => {
    let fileURL = this.props.officeContext.document.url;
    if (fileURL.indexOf('sharepoint.com') > -1) {
      // let docName = fileURL.split('/')[fileURL.split('/').length - 1];
      let serverRelativeUrl = fileURL.split(appConst.webUrl)[fileURL.split(appConst.webUrl).length - 1];

      let response: any = await GetSPDocSSO(serverRelativeUrl, {});
      // console.log(response);
      let docGUID = "";
      if (!response) {
        throw new Error("Middle tier didn't respond");
      } else {
        // this.setState({ docGUID: JSON.parse(response).d.UniqueId }, () => { this.GetSPData() });
        docGUID = JSON.parse(response).d.UniqueId;
      }
      return docGUID;
    }
  }
  public checkUser = async () => {
    let response: any = await checkGroup({});
    console.log(response);
    let isGroup = [];
    if (response && response.value.length > 0)
      isGroup = response.value.filter((res) => { return res.displayName == 'CoCoInitiatorGrp' });
    isGroup.length > 0 ? console.log('Access Granted') : console.log('Access Denied');
  }
  public currentUserDetails = async () => {
    let response: any = await currentuserDetails({});
    console.log(response);
    if (response)
      return response.mail;
    else
      return "";
  }
  public render(): ReactElement<IProps> {
    return (
      <div style={{ backgroundColor: "lightgrey" }}>
        <Pivot>
          <PivotItem headerText="Section Content">
            {/* <div className="UserDashboard">"Tab1"</div> */}
            <FullFormComponent officeContext={this.props.officeContext} sectionInfo={this.state.mappedSectionInfo} docGUID={this.state.docGUID} docInfo={this.state.docInfo} />
          </PivotItem>
          <PivotItem linkText="Document Status">
            {/*  <div className="UserDashboard">"Tab2"</div> */}
            <SectionDetails
              sectionInfo={this.state.SectionsDetails}
              documentID={this.state.docGUID}
              currentUserEmail={this.state.currentUserEmail}
            />
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}

export default HomeScreenComponent;

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
import * as appConst from "../../constants/appConst";


export interface IProps extends ComboboxProps {
  officeContext: any;
}

export interface IState {
  docGUID: string;
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
    };
  }
  async componentDidMount(): Promise<void> {
    let tempState = { ...this.state };
    const currEmail = await this.currentUserDetails();
    tempState = { ...tempState, ...{ currentUserEmail: currEmail } };
    // await this.checkUser();
    const docGUID = await this.getDocumentMetadata();
    tempState = { ...tempState, ...{ docGUID: docGUID } };
    const docInfo = await this.GetSPDocDetails(docGUID);
    if (docInfo) {
      tempState = { ...tempState, ...{ docInfo: docInfo } };
      const sectionInfo = await this.GetSPAssigneeData(docGUID);
      const mappedSectionInfo: SectionAssignment[] = this.mapSectionItemToRow(sectionInfo, docGUID)
      const sortedSectionInfo = mappedSectionInfo.sort((a, b) => a.SectionNumber - b.SectionNumber)
      tempState = { ...tempState, ...{ SectionsDetails: sectionInfo, mappedSectionInfo: sortedSectionInfo } };
    }
    this.setState(tempState);
  }
  public updateDocInfo=(docInfoObj)=>{
    this.setState({docInfo:{...this.state.docInfo,...docInfoObj}});
  }
  public mapSectionItemToRow = (sectionInfo, docGUID) => {
    return sectionInfo.map((item, index) => {
      const contributors = item.Contributors.results && item.Contributors.results.length > 0 ? item.Contributors.results.map(item => {
        return {
          ContributorID: item.Id,
          ContributorDisplayName: item.Title
        }
      }) : [];

      const sectionAssignment: SectionAssignment = {
        itemID: item.Id,
        SectionNumber: item.SectionSequence,
        SectionName: item.SectionName,
        POwnerID: item.PrimaryOwnerId ? item.PrimaryOwnerId : 0,
        POwnerDisplayName: item.PrimaryOwner.Title,
        POwnerEmail: "",
        SOwnerID: item.SecondaryOwnerId ? item.SecondaryOwnerId : 0,
        SOwnerDisplayName: item.SecondaryOwner.Title,
        SOwnerEmail: "",
        Contributors: contributors,
        DeadLineDate: item.TargetDate ? new Date(item.TargetDate) : null,
        DocumentID: item.DocumentID ? item.DocumentID : docGUID,
        SectionID: item.SectionID ? item.SectionID : ""
      }
      return sectionAssignment;
    })
  }
  public GetSPDocDetails = async (docGUID) => {
    const filter = `DocumentID eq '${docGUID}'`;
    let response = await GetSPListData(appConst.lists.documentDetails, "*", "", filter);
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
    let response: any = await GetSPListData(appConst.lists.assigneeDetails, "*,PrimaryOwner/Title,SecondaryOwner/Title,Contributors/Id,Contributors/Title", "PrimaryOwner,SecondaryOwner,Contributors", filter);
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
      let serverRelativeUrl = fileURL.split(appConst.webUrl)[fileURL.split(appConst.webUrl).length - 1];

      let response: any = await GetSPDocSSO(serverRelativeUrl, {});
      let docGUID = "";
      if (!response) {
        throw new Error("Middle tier didn't respond");
      } else {
        docGUID = JSON.parse(response).d.UniqueId;
      }
      return docGUID;
    }
  }
  // public checkUser = async () => {
  //   let response: any = await checkGroup({});
  //   console.log(response);
  //   let isGroup = [];
  //   if (response && response.value.length > 0)
  //     isGroup = response.value.filter((res) => { return res.displayName == 'CoCoInitiatorGrp' });
  //   isGroup.length > 0 ? console.log('Access Granted') : console.log('Access Denied');
  // }
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
            <FullFormComponent officeContext={this.props.officeContext} sectionInfo={this.state.mappedSectionInfo} docGUID={this.state.docGUID} docInfo={this.state.docInfo} updateDocInfo={this.updateDocInfo} />
          </PivotItem>
          <PivotItem linkText="Document Status">
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

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
import { GetSPDocSSO, getSearchUser } from "../../helpers/sso-helper";
import { DatePicker, DayOfWeek, DefaultButton, Pivot, PivotItem, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import SectionFormComponent from "./SectionForm";
import FullFormComponent from "./FullForm";

const MyComponent = () => {
  return <div></div>;
};

export interface IProps extends ComboboxProps {
  officeContext: any;
}

export interface IState {
  docGUID:string;
  // options: any[];
  // Sections: SectionAssignment[];
}

export class HomeScreenComponent extends Component<IProps, IState> {
  constructor(props: IProps) {
    super(props);
    this.state = {
      docGUID:""
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
    await this.getDocumentMetadata();
  }
  public getDocumentMetadata = async () => {
    let fileURL = this.props.officeContext.document.url;
    if (fileURL.indexOf('sharepoint.com') > -1) {
      // let docName = fileURL.split('/')[fileURL.split('/').length - 1];
      let serverRelativeUrl = fileURL.split('https://vichitra.sharepoint.com')[fileURL.split('https://vichitra.sharepoint.com').length - 1];

      let response: any = await GetSPDocSSO(serverRelativeUrl, {});
      console.log(response);
      if (!response) {
        throw new Error("Middle tier didn't respond");
      } else {
        this.setState({ docGUID: JSON.parse(response).d.UniqueId });
      }
    }
  }

  public render(): ReactElement<IProps> {
    return (
      <div style={{ backgroundColor: "lightgrey" }}>
        <Pivot>
          <PivotItem headerText="In Progress">
            <FullFormComponent officeContext={this.props.officeContext} docGUID={this.state.docGUID}/>
          </PivotItem>
          <PivotItem linkText="Completed">
            <div className="UserDashboard">"Tab2"</div>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}

export default HomeScreenComponent;

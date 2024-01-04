import React, { Component, ReactElement } from "react";
import { Combobox, Label, Option, Persona, useId, } from "@fluentui/react-components";
import type { ComboboxProps } from "@fluentui/react-components";
import { GetSiteUserFromEmailSSO, getSearchUser } from "../../helpers/sso-helper";
import { IComboBox, IComboBoxOption } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";

const MyComponent = () => {
  return <div></div>;
};

export interface IProps extends ComboboxProps {
  isReadOnlyForm: boolean;
  fieldState: string;
  fieldName: string;
  sectionInfo: any;
  isMandatory: boolean;
  sectionNumber: number;
  forceResetGrid?: boolean;
  updatePeopleCB: Function;
  allADUsers: any[];
}

export interface IState {
  options: any[];
  pplValue: string;
  selectedOptions: any[];
}

const options = [
  { displayName: "Jatin Batra" },
  { displayName: "Karan Aggarwal" },
  { displayName: "Harsha Reddy" },
  { displayName: "Shivam Rastogi" },
];

export class ShowADUserComponent extends Component<IProps, IState> {
  constructor(props: IProps) {
    super(props);

    this.state = {
      options: this.props.allADUsers,
      pplValue: this.props.fieldState == "Contributor" ? "" : this.props.sectionInfo[`${this.props.fieldState}DisplayName`],
      selectedOptions: this.props.fieldState == "Contributor" ? [] : [this.props.sectionInfo[`${this.props.fieldState}ID`]]
    };
  }
  public componentDidUpdate(prevProps: Readonly<IProps>, prevState: Readonly<IState>, snapshot?: any): void {
    if ((prevProps.forceResetGrid != this.props.forceResetGrid) && this.props.forceResetGrid) {
      this.setState({
        pplValue: this.props.fieldState == "Contributor" || this.props.forceResetGrid ? "" : this.props.sectionInfo[`${this.props.fieldState}DisplayName`],
        selectedOptions: this.props.fieldState == "Contributor" || this.props.forceResetGrid ? [] : [this.props.sectionInfo[`${this.props.fieldState}ID`]]
      });
    }
  }
  private onComboBoxChange: ComboboxProps["onChange"] = (event) => {
    const changedText = event.target.value.trim();
    if (changedText.length > 2) {
      const searchedUsers = this.props.allADUsers.filter((item) => {
        if (item.displayName.indexOf(changedText) == 0) {
          return item;
        }
      });
      console.log(searchedUsers);
      this.setState({ options: searchedUsers, pplValue: changedText });
    } else {
      this.setState({ options: [], pplValue: changedText });
    }
  };
  private oncomboBoxSelect = (data) => {
    if (data.selectedOptions.length > 0) {


      if (this.props.fieldState == "Contributor") {
        this.setState({ options: this.props.allADUsers, selectedOptions: [], pplValue: "" }, () => {
          this.props.updatePeopleCB({
            ContributorDisplayName: data.optionText,
            ContributorID: data.optionValue
          });
        });
      }
      else {
        this.setState({ selectedOptions: [data.optionValue], pplValue: data.optionText }, () => {
          let updatedObj = {};
          updatedObj[`${this.props.fieldState}DisplayName`] = data.optionText;
          updatedObj[`${this.props.fieldState}ID`] = data.optionValue;
          this.props.updatePeopleCB(this.props.sectionNumber, updatedObj);
        });
      }
    }

  }

  public render(): ReactElement<IProps> {
    return (
      <div>
        {this.props.fieldName != "" && <Label>{this.props.fieldName}</Label>
        }
        {this.props.fieldName != "" && this.props.isMandatory && <span style={{ color: "red" }}>*</span>}

        <Combobox
          disabled={this.props.isReadOnlyForm}
          freeform
          onChange={this.onComboBoxChange}
          style={{ width: "100%" }}
          onOptionSelect={(event, data) => this.oncomboBoxSelect(data)}
          value={this.state.pplValue}
          selectedOptions={this.state.selectedOptions}
        >
          {this.state.options.map((option, index) => (
            <Option text={option.displayName} value={option.SPUSerId} key={index}>
              <Persona avatar={{ color: "colorful", "aria-hidden": true }} name={option.displayName} />
            </Option>
          ))}
        </Combobox>
      </div>
    );
  }
}

export default ShowADUserComponent;

import React, { Component, ReactElement } from "react";
import { Combobox, Label, Option, Persona, useId, } from "@fluentui/react-components";
import type { ComboboxProps } from "@fluentui/react-components";
import { GetSiteUserFromEmailSSO, getSearchUser } from "../../helpers/sso-helper";
import { IComboBox, IComboBoxOption } from "@fluentui/react";

const MyComponent = () => {
  return <div></div>;
};

export interface IProps extends ComboboxProps {
  fieldState: string;
  fieldName: string;
  sectionInfo:any;
  isMandatory: boolean;
  sectionNumber: number;
  updatePeopleCB: Function;
  allADUsers: any[];
}

export interface IState {
  options: any[];
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
      options: [],
    };
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
      this.setState({ options: searchedUsers });
      // getSearchUser(changedText, (result: any) => {
      //   this.setState({ options: result.value });
      //   console.log(result);
      // });
    } else {
      this.setState({ options: [] });
    }
    // const matches = options.filter((option) => option.toLowerCase().indexOf(value.toLowerCase()) === 0);
    // setMatchingOptions(matches);
  };
  private oncomboBoxSelect = (data) => {
    console.log(data);
    let updatedObj = {};
    updatedObj[`${this.props.fieldState}Email`] = data.optionText;    
    updatedObj[`${this.props.fieldState}ID`] = data.optionValue;
    this.props.updatePeopleCB(this.props.sectionNumber, updatedObj);
    // GetSiteUserFromEmailSSO(data.optionValue, (result: any) => {
    //   console.log(result);
    //   updatedObj[`${this.props.fieldState}ID`] = JSON.parse(result)?.d?.Id;
    //   this.props.updatePeopleCB(this.props.sectionNumber, updatedObj);
    // });
    //JSON.parse(response).d.Id
  }

  public render(): ReactElement<IProps> {
    return (
      <div>
        {this.props.fieldName != "" && <Label>{this.props.fieldName}</Label>
        }
        {this.props.fieldName != "" && this.props.isMandatory && <span style={{ color: "red" }}>*</span>}
        <Combobox
          freeform
          onChange={this.onComboBoxChange}
          style={{ width: "100%" }}
          // required={this.props.isMandatory}
          onOptionSelect={(event, data) => this.oncomboBoxSelect(data)}
          selectedOptions={[this.props.sectionInfo[`${this.props.fieldState}ID`]]}
        >
          {this.state.options.map((option) => (
            <Option text={option.displayName} value={option.SPUSerId}>
              <Persona avatar={{ color: "colorful", "aria-hidden": true }} name={option.displayName} />
            </Option>
          ))}
        </Combobox>
      </div>
    );
  }
}

export default ShowADUserComponent;

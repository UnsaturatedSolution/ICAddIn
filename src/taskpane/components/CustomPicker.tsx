import React, { Component, ReactElement } from "react";
import { Combobox, Label, Option, Persona, useId } from "@fluentui/react-components";
import type { ComboboxProps } from "@fluentui/react-components";
import { getSearchUser } from "../../helpers/sso-helper";

const MyComponent = () => {
  return <div></div>;
};

export interface IProps extends ComboboxProps {
  fieldID: string;
  fieldName: string;
  isMandatory: boolean;
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
      getSearchUser(changedText, (result: any) => {
        this.setState({ options: result.value });
      });
    } else {
      this.setState({ options: [] });
    }
    // const matches = options.filter((option) => option.toLowerCase().indexOf(value.toLowerCase()) === 0);
    // setMatchingOptions(matches);
  };

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
        >
          {this.state.options.map((option) => (
            <Option text={option.displayName}>
              <Persona avatar={{ color: "colorful", "aria-hidden": true }} name={option.displayName} />
            </Option>
          ))}
        </Combobox>
      </div>
    );
  }
}

export default ShowADUserComponent;

import React, { Component, ReactElement } from "react";
import {
    Combobox,
    Label,
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
import { getSearchUser } from "../../helpers/sso-helper";
import { DatePicker, DayOfWeek, DefaultButton, Pivot, PivotItem, defaultDatePickerStrings } from "@fluentui/react";
import { SectionAssignment } from "./ISectionAssignment";
import ShowADUserComponent from "./CustomPicker";

const MyComponent = () => {
    return <div></div>;
};

export interface IProps extends ComboboxProps {
    sectionInfo: SectionAssignment;
    updateParentSectionState: Function;
}

// export interface IState {
//   options: any[];
//   Sections: SectionAssignment[];
// }

export class SectionFormComponent extends Component<IProps> {
    constructor(props: IProps) {
        super(props);
        // this.state = {
        //   options: [],
        //   Sections: [],
        // };
    }

   
    public updateParentSectionState = () => {
        let tempSectionItem = { ...this.props.sectionInfo };
        this.props.updateParentSectionState(tempSectionItem);
    }
    public render(): ReactElement<IProps> {
        return (
            <div style={{ borderTop: "solid darkgrey 1px" }}>
                <h3>{`Section ${this.props.sectionInfo.SectionNumber}`}</h3>
                {/* <div>
                    <ShowADUserComponent fieldName="Primary Owner" isMandatory={true}></ShowADUserComponent>
                </div>
                <div>
                    <ShowADUserComponent fieldName="Secondary Owner" isMandatory={true}></ShowADUserComponent>
                </div>
                <div>
                    <ShowADUserComponent fieldName="Contributor(s)" isMandatory={false}></ShowADUserComponent>
                </div> */}
                <div>
                    <DatePicker
                        label={"Target Date"}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        onSelectDate={(date) => {
                            let tempSectionItem = { ...this.props.sectionInfo };
                            tempSectionItem.DeadLineDate = date;
                        }}
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
                    />
                </div>
            </div>
        );
    }
}

export default SectionFormComponent;

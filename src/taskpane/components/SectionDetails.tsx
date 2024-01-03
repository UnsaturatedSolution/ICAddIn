import React, { Component, ReactElement } from "react";
import type { ComboboxProps } from "@fluentui/react-components";
import { DefaultButton, DetailsList, IColumn } from "@fluentui/react";
import { UpdateRequestSSO, GetSPListData } from "../../helpers/sso-helper";
export interface IProps extends ComboboxProps {
    sectionInfo: any[];
    documentID: string;
    currentUserEmail: string;
}

const listColumns: IColumn[] = [
    { key: 'SectionSequence', name: 'Sr No', fieldName: 'SectionSequence', minWidth: 50, maxWidth: 50, isResizable: true },
    {
        key: 'SectionName', name: 'Section', fieldName: 'SectionName', minWidth: 60, maxWidth: 60, isResizable: true
    },
    {
        key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 80, isResizable: true
    },
    {
        key: 'PrimaryOwnerStringId', name: 'PrimaryOwner', minWidth: 100, maxWidth: 100, isResizable: true,
    },
    {
        key: 'SecondaryOwnerStringId', name: 'SecondaryOwner', minWidth: 110, maxWidth: 110, isResizable: true,
    },
    {
        key: 'TargetDate', name: 'Target Date', fieldName: 'TargetDate', minWidth: 100, maxWidth: 100, isResizable: true,
    },
];
export class SectionDetails extends Component<IProps> {
    constructor(props: IProps) {
        super(props);
    }

    private statusDisplay = (status: string) => {
        switch (status) {
            case "InProgress": return <span style={{ backgroundColor: "rgb(255, 209, 0)", color: "white", fontWeight: "bold" }}>{status}</span>;
            case "Completed": return <span style={{ backgroundColor: "rgb(0, 204, 0)", color: "white", fontWeight: "bold" }}>{status}</span>;
            case "NotStarted": return <span style={{ backgroundColor: "rgb(166, 166, 166)", color: "white", fontWeight: "bold" }}>{status}</span>;

        }
    }
    private onRenderItem = (item?: any, index?: number, column?: IColumn) => {
        switch (column.name) {
            case "PrimaryOwner":
                return item.PrimaryOwnerId != null && item.PrimaryOwnerId != undefined && item.PrimaryOwnerId != '' ? item.PrimaryOwner.Title : '';
            case "SecondaryOwner":
                return item.SecondaryOwnerId != null && item.SecondaryOwnerId != undefined && item.SecondaryOwnerId != '' ? item.SecondaryOwner.Title : '';
            case "Status":
                return this.statusDisplay(item.Status);
            default:
                return item[column.fieldName];
        }
    }
    private sendMail = async () => {
        const query = `$select=*&$filter=DocumentID eq '${this.props.documentID}'&$top=1&$orderby=Modified desc`
        const documentDetails = await GetSPListData("InvestCorpDocumentDetails", "*", "", query);
        const result = JSON.parse(documentDetails);
        const ItemID = documentDetails ? result.d.results[0]["Id"] : 0;
        const updateDetails = { SendReport: true, SendReportToEmailAddr: this.props.currentUserEmail }
        await UpdateRequestSSO(updateDetails, ItemID, 'InvestCorpDocumentDetails');
    }
    public render(): ReactElement<IProps> {
        return (
            <div className={`ms-Grid-row`} >
                {/* <div className={`ms-Grid-col ms-sm12`} style={{ display: "flex", alignItems: 'center' }}>
                 */}    <DetailsList
                    columns={listColumns}
                    items={this.props.sectionInfo}
                    //items={this.state.approvalHistoryItems} 
                    onRenderItemColumn={this.onRenderItem}
                />
                <DefaultButton
                    style={{ color: "#000", backgroundColor: "white", alignItems: "left" }}
                    text="Mail Report"
                    iconProps={{ iconName: "MailAttached" }}
                    onClick={this.sendMail}
                />
                {/* </div> */}
            </div >
        );
    }
}

export default SectionDetails;
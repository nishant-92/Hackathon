import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import "./SprintPlanner.scss";

import { Button } from "azure-devops-ui/Button";
import { ObservableArray, ObservableValue } from "azure-devops-ui/Core/Observable";
import { localeIgnoreCaseComparer } from "azure-devops-ui/Core/Util/String";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { ListSelection, ISimpleListCell } from "azure-devops-ui/List";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { Header } from "azure-devops-ui/Header";
import { Page } from "azure-devops-ui/Page";
import { TextField } from "azure-devops-ui/TextField";
import { Dialog } from "azure-devops-ui/Dialog";
import { Observer } from "azure-devops-ui/Observer";
import { Card } from "azure-devops-ui/Card";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { ColumnFill, ColumnMore, renderSimpleCell, Table } from "azure-devops-ui/Table";
import { ArrayItemProvider, getItemsValue } from "azure-devops-ui/Utilities/Provider";

import { CommonServiceIds, getClient, IProjectPageService } from "azure-devops-extension-api";
import { IWorkItemFormNavigationService, WorkItemTrackingRestClient, WorkItemTrackingServiceIds,Wiql } from "azure-devops-extension-api/WorkItemTracking";

import { showRootComponent } from "../../Common";

type callback = (status:boolean) => void;
type PreviewProp = {
    status: boolean;
    selectedDuration: string[],
    selectedMembers: string[],
    callback:callback;
}

type PreviewState = {
    sprintDetails : ArrayItemProvider<ITableItem>;
    status : boolean;
}

export default class SprintPreviewContent extends React.Component<PreviewProp, {}> {
    private isDialogOpen : boolean;
    private callback: callback;
    private selection = new ListSelection();

    constructor(props: PreviewProp) {
        super(props);
        this.isDialogOpen = props.status;
        this.callback = props.callback;
    }

    state: PreviewState ={
        status:false,
        sprintDetails:new ArrayItemProvider<ITableItem>([])
    }

    public componentDidMount() {
        SDK.init();
        this.loadData();
    }
    
    public render(): JSX.Element {
        const onDismiss = () => {
            this.isDialogOpen = false;
            this.callback(false);
        };
        console.log("Inside sprint preview");
        console.log("Sprint Duration:" + this.props.selectedDuration);
        console.log("Selected Team Members:" + this.props.selectedMembers);
        return (
            <Page className="sample-hub flex-grow">
                <Header title="Sprint Planner" />
                <div className="page-content">
                    <div className="sample-form-section flex-row flex-center">
                        <Dialog
                            titleProps={{ text: "Sprint Preview" }}
                            footerButtonProps={[
                                {
                                    text: "Close",
                                    onClick: onDismiss
                                }
                            ]}
                            onDismiss={onDismiss}
                            calloutContentClassName = "dialogbox-size"
                        >
                            {this.state.status && <Table<Partial<ITableItem>>
                                ariaLabel="Sprint Preview"
                                columns={sizableColumns}
                                itemProvider={this.state.sprintDetails}
                                selection={this.selection}
                                onSelect={(event, data) => console.log("Select Row - " + data.index)}
                            />}
                        </Dialog>
                    </div>
                </div>
            </Page>
        );
    }

    private loadData():void{
        var promise = this.allWorkItems();
        promise.then((data) => {this.setState({
            status:true,
            sprintDetails:data
        })})
    }

    private async allWorkItems() {
        const client = getClient(WorkItemTrackingRestClient);
        const model: Wiql = {
            query : "select * From WorkItems where [State] = 'Active' OR [State] = 'New' Order By [System.AssignedTo] Asc, [id] Asc"
        }
        const types = await client.queryByWiql(model,"MSHackathon2020","MSHackathon2020",false,100);
        var workItemsResult = types.workItems;
        var workItemArray = new ArrayItemProvider<ITableItem>([]);
        
        var map = new Map();
        map.set(2,4);map.set(3,5);map.set(5,3);map.set(8,3);
        map.set(9,7);map.set(10,8);map.set(11,2);map.set(12,3);
        var sprintDuration = parseInt(this.props.selectedDuration[0][0]);
        for(var i = 0; i<workItemsResult.length;i++)
        {
            const types = await client.getWorkItem(workItemsResult[i].id,"MSHackathon2020",['System.Title','System.AssignedTo','Microsoft.VSTS.Common.Priority'],new Date(),0);

            if(this.props.selectedMembers.find(Element =>  Element == types.fields["System.AssignedTo"]["displayName"]))
            {
                var workItem : ITableItem = {
                    workItem: { iconProps: { iconName: types.fields["System.AssignedTo"]["_links"]["avatar"]["href"] }, text: types.fields["System.Title"] },
                    wid: workItemsResult[i].id,
                    assignedTo: types.fields["System.AssignedTo"]["displayName"],
                    estimate: map.get(workItemsResult[i].id),
                    priority: types.fields["Microsoft.VSTS.Common.Priority"],
                }
                workItemArray.value.push(workItem);
            }
        }
        return workItemArray;
    }
}

interface ITableItem {
    workItem: ISimpleListCell;
    wid: number;
    assignedTo: string;
    estimate: number;
    priority: number;
}

function onSizeSizable(event: MouseEvent, index: number, width: number) {
    (sizableColumns[index].width as ObservableValue<number>).value = width;
}

const sizableColumns = [
    {
        id: "workItem",
        name: "WorkItem",
        minWidth: 50,
        width: new ObservableValue(300),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    {
        id: "wid",
        name: "Work Item ID",
        maxWidth: 300,
        width: new ObservableValue(200),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    {
        id: "assignedTo",
        name: "AssignedTo",
        maxWidth: 300,
        width: new ObservableValue(200),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    {
        id: "estimate",
        name: "Estimate",
        maxWidth: 300,
        width: new ObservableValue(200),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    { id: "priority", name: "Priority", width: new ObservableValue(200), renderCell: renderSimpleCell },
    ColumnFill
];

function onSizeMore(event: MouseEvent, index: number, width: number) {
    (moreColumns[index].width as ObservableValue<number>).value = width;
}

const moreColumns = [
    {
        id: "workItem",
        minWidth: 50,
        name: "WorkItem",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(200)
    },
    {
        id: "wid",
        maxWidth: 300,
        name: "Work Item ID",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(100)
    },
    {
        id: "assignedTo",
        maxWidth: 300,
        name: "AssignedTo",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(100)
    },
    {
        id: "estimate",
        name: "Estimate",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(100)
    },
    {
        id: "priority",
        name: "Priority",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(100)
    },
    ColumnFill,
    new ColumnMore(() => {
        return {
            id: "sub-menu",
            items: [
                { id: "submenu-one", text: "SubMenuItem 1" },
                { id: "submenu-two", text: "SubMenuItem 2" }
            ]
        };
    })
];

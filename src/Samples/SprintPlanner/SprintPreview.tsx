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


    private loadTableData():PreviewState{

        return {
            status:true,
            sprintDetails : new ArrayItemProvider<ITableItem>([
                {
                    workItem: { iconProps: { iconName: "Home" }, text: "Test task 1" },
                    assignedTo: "Ankit",
                    estimate: 2
                },
                {
                    workItem: { iconProps: { iconName: "Home" }, text: "Test task 2" },
                    assignedTo: "Nishant",
                    estimate: 1
                },
                {
                    workItem: { iconProps: { iconName: "Home" }, text: "Test task 3" },
                    assignedTo: "Ashish",
                    estimate: 3
                }
            ]),
        }
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
            query : "select * From WorkItems where [State] = 'Active' OR [State] = 'New' "
        }
        const types = await client.queryByWiql(model,"MSHackathon2020","MSHackathon2020",false,100);
        var workItemsResult = types.workItems;
        var workItemArray = new ArrayItemProvider<ITableItem>([]);
 
        for(var i = 0; i<workItemsResult.length;i++)
        {
            console.log(workItemsResult[i].id);
            const types = await client.getWorkItem(workItemsResult[i].id,"MSHackathon2020",['System.Title','System.AssignedTo','Microsoft.VSTS.Common.Priority'],new Date(),0);
            
            var workItem : ITableItem = {
                workItem: { iconProps: { iconName: types.fields["System.AssignedTo"]["_links"]["avatar"]["href"] }, text: types.fields["System.Title"] },
                assignedTo: types.fields["System.AssignedTo"]["displayName"],
                estimate: types.fields["Microsoft.VSTS.Common.Priority"]
            }
            workItemArray.value.push(workItem);
        }
        return workItemArray;
    }

}

interface ITableItem {
    workItem: ISimpleListCell;
    assignedTo: string;
    estimate: number;
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
        id: "assignedTo",
        name: "AssignedTo",
        maxWidth: 300,
        width: new ObservableValue(200),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    { id: "estimate", name: "Estimate", width: new ObservableValue(200), renderCell: renderSimpleCell },
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

const tableItems = new ArrayItemProvider<ITableItem>([
    {
        workItem: { iconProps: { iconName: "Home" }, text: "Test task 1" },
        assignedTo: "Ankit",
        estimate: 2
    },
    {
        workItem: { iconProps: { iconName: "Home" }, text: "Test task 2" },
        assignedTo: "Nishant",
        estimate: 1
    },
    {
        workItem: { iconProps: { iconName: "Home" }, text: "Test task 3" },
        assignedTo: "Ashish",
        estimate: 3
    }
]);

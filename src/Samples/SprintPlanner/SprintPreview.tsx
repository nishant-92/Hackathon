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
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

import { CommonServiceIds, getClient, IProjectPageService } from "azure-devops-extension-api";
import { IWorkItemFormNavigationService, WorkItemTrackingRestClient, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";

import { showRootComponent } from "../../Common";

type callback = (status:boolean) => void;
type PreviewProp = {
    status: boolean;
    callback:callback;
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
    
    public componentDidMount() {
        SDK.init();
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
                            titleProps={{ text: "Confirm" }}
                            footerButtonProps={[
                                {
                                    text: "Close",
                                    onClick: onDismiss
                                }
                            ]}
                            onDismiss={onDismiss}
                        >
                            <Table<Partial<ITableItem>>
                                ariaLabel="Food Inventory Table"
                                columns={sizableColumns}
                                itemProvider={tableItems}
                                selection={this.selection}
                                onSelect={(event, data) => console.log("Select Row - " + data.index)}
                            />
                        </Dialog>
                    </div>
                </div>
            </Page>
        );
    }

}

interface ITableItem {
    name: ISimpleListCell;
    calories?: number;
    cost?: string;
}

function onSizeSizable(event: MouseEvent, index: number, width: number) {
    (sizableColumns[index].width as ObservableValue<number>).value = width;
}

const sizableColumns = [
    {
        id: "name",
        name: "Name",
        minWidth: 50,
        width: new ObservableValue(300),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    {
        id: "calories",
        name: "Calories",
        maxWidth: 300,
        width: new ObservableValue(200),
        renderCell: renderSimpleCell,
        onSize: onSizeSizable
    },
    { id: "cost", name: "Cost", width: new ObservableValue(200), renderCell: renderSimpleCell },
    ColumnFill
];

function onSizeMore(event: MouseEvent, index: number, width: number) {
    (moreColumns[index].width as ObservableValue<number>).value = width;
}

const moreColumns = [
    {
        id: "name",
        minWidth: 50,
        name: "Name",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(200)
    },
    {
        id: "calories",
        maxWidth: 300,
        name: "Calories",
        onSize: onSizeMore,
        readonly: true,
        renderCell: renderSimpleCell,
        width: new ObservableValue(100)
    },
    {
        id: "cost",
        name: "Cost",
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
        name: { iconProps: { iconName: "Home" }, text: "Potato Chips" },
        calories: 400,
        cost: "$2.99"
    },
    {
        name: { iconProps: { iconName: "Home" }, text: "Yogurt" },
        calories: 140,
        cost: "$3.99"
    },
    {
        name: { iconProps: { iconName: "Home" }, text: "Milk" },
        cost: "$2.50"
    }
]);

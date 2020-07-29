import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import "./SprintPlanner.scss";

import { Button } from "azure-devops-ui/Button";
import { ObservableArray, ObservableValue } from "azure-devops-ui/Core/Observable";
import { localeIgnoreCaseComparer } from "azure-devops-ui/Core/Util/String";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { Checkbox } from "azure-devops-ui/Checkbox";
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
import { CoreRestClient } from "azure-devops-extension-api/Core";

const firstCheckbox = new ObservableValue<boolean>(false);
const secondCheckbox = new ObservableValue<boolean>(false);
const thirdCheckbox = new ObservableValue<boolean>(false);

class SprintPlannerContent extends React.Component<{}, {}> {

    private workItemIdValue = new ObservableValue("1");
    private workItemTypeValue = new ObservableValue("Bug");
    private selection = new ListSelection();
    private workItemTypes = new ObservableArray<IListBoxItem<string>>();
    private teamMembers = new ObservableArray<string>();
    private isDialogOpen = new ObservableValue<boolean>(false);

    constructor(props: {}) {
        super(props);
    }

    public componentDidMount() {
        SDK.init();
        this.loadWorkItemTypes();
        this.loadAccounts();
    }

    public render(): JSX.Element {
        const onDismiss = () => {
            this.isDialogOpen.value = false;
        };
        return (
            <Page className="sample-hub flex-grow">
                <Header title="Sprint Planner" />
                <div className="page-content">
                    <div className="sample-form-section flex-row flex-center">
                        <TextField className="sample-work-item-id-input" label="Sprint Duration" value={this.workItemIdValue} onChange={(ev, newValue) => { this.workItemIdValue.value = newValue; }} />
                        <Button className="sample-work-item-button" text="Open..." onClick={() => this.onOpenExistingWorkItemClick()} />
                    </div>
                    <div className="sample-form-section flex-row flex-center">
                        <div className="flex-column">
                            <label htmlFor="work-item-type-picker">New work item type:</label>
                            <Dropdown<string>
                                className="sample-work-item-type-picker"
                                items={this.workItemTypes}
                                onSelect={(event, item) => { this.workItemTypeValue.value = item.data! }}
                                selection={this.selection}
                            />
                        </div>
                        <Button className="sample-work-item-button" text="Sprint Preview" onClick={() => {this.isDialogOpen.value = true;}} />
                         
                            <Observer isDialogOpen={this.isDialogOpen}>
                            {(props: { isDialogOpen: boolean }) => {
                                return props.isDialogOpen ? (
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
                                ) : null;
                            }}
                        </Observer>

                    </div>
                    <div className="sample-form-section flex-row flex-center">
                        <div className="flex-column">
                            <label htmlFor="account-picker">Team members:</label>
                            <Observer checkboxes={this.teamMembers}>
                                {
                                    (props: {checkboxes: Array<string>}) => {
                                        console.log(props.checkboxes, this.teamMembers);
                                        console.log(213);
                                        return props.checkboxes.length ? (
                                            <div>
                                                <Checkbox className="sample-account-picker"
                                                    onChange={(event, checked) => (firstCheckbox.value = checked)}
                                                    checked={firstCheckbox}
                                                    label={props.checkboxes[0]}
                                                />
                                                <Checkbox className="sample-account-picker"
                                                    onChange={(event, checked) => (secondCheckbox.value = checked)}
                                                    checked={secondCheckbox}
                                                    label={props.checkboxes[1]}
                                                />
                                                <Checkbox className="sample-account-picker"
                                                    onChange={(event, checked) => (thirdCheckbox.value = checked)}
                                                    checked={thirdCheckbox}
                                                    label={props.checkboxes[2]}
                                                />
                                                </div>
                                        ) : (
                                            <div>loading</div>
                                        );
                                    }
                                }
                            </Observer>
                        </div>
                        <Button className="sample-work-item-button" text="New..." onClick={() => this.onOpenNewWorkItemClick()} />
                    </div>
                </div>
            </Page>
        );
    }

    private async loadAccounts(): Promise<void> {

        console.log("Entering load accounts");
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        console.log("Get peoject service");
        const project = await projectService.getProject();
        console.log("Get project");

        let teamMemberNames: string[] = [];

        if (project)  {
            const client = await getClient(CoreRestClient);
            console.log("Get core client");
            const teamMembers = await client.getTeamMembersWithExtendedProperties(project.name, "MSHackathon2020");
            console.log("Get team members");
            teamMemberNames = teamMembers.map(t => t.identity.displayName);
            teamMemberNames.sort((a, b) => localeIgnoreCaseComparer(a, b));

            console.log("team members:" + teamMemberNames);
        }

        this.teamMembers.change(0, ...teamMemberNames);
        console.log("all team member:" + this.teamMembers);
        console.log("1st team member:" + this.teamMembers.value[0]);
        console.log("exit team");
        //this.selection.select(0);
    }

    private async loadWorkItemTypes(): Promise<void> {

        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        let workItemTypeNames: string[];

        if (!project) {
            workItemTypeNames = [ "Issue" ];
        }
        else {
            const client = getClient(WorkItemTrackingRestClient);
            const types = await client.getWorkItemTypes(project.name);
            workItemTypeNames = types.map(t => t.name);
            workItemTypeNames.sort((a, b) => localeIgnoreCaseComparer(a, b));
        }

        this.workItemTypes.push(...workItemTypeNames.map(t => { return { id: t, data: t, text: t } }));
        this.selection.select(0);
    }

    private async onOpenExistingWorkItemClick() {
        const navSvc = await SDK.getService<IWorkItemFormNavigationService>(WorkItemTrackingServiceIds.WorkItemFormNavigationService);
        navSvc.openWorkItem(parseInt(this.workItemIdValue.value));
    };

    private async onOpenNewWorkItemClick() {
        const navSvc = await SDK.getService<IWorkItemFormNavigationService>(WorkItemTrackingServiceIds.WorkItemFormNavigationService);
        navSvc.openNewWorkItem(this.workItemTypeValue.value, { 
            Title: "Opened a work item from the Work Item Nav Service",
            Tags: "extension;wit-service",
            priority: 1,
            "System.AssignedTo": SDK.getUser().name,
         });
    };
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

showRootComponent(<SprintPlannerContent />);
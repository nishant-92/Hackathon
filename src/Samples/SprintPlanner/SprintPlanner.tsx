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
import { Observer } from "azure-devops-ui/Observer";
import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";

import { CommonServiceIds, getClient, IProjectPageService } from "azure-devops-extension-api";
import { showRootComponent } from "../../Common";
import { CoreRestClient } from "azure-devops-extension-api/Core";
import SprintPreviewContent from "./SprintPreview"

const firstCheckbox = new ObservableValue<boolean>(false);
const secondCheckbox = new ObservableValue<boolean>(false);
const thirdCheckbox = new ObservableValue<boolean>(false);


type SprinPlannerState = {
    showMessage: boolean,
    selectedDuration: string[],
    selectedMembers: string[],
};

class SprintPlannerContent extends React.Component<{}, {}> {

    private workItemTypeValue = new ObservableValue("Bug");
    private selection = new ListSelection();
    private teamMembers = new ObservableArray<string>();
    private isDialogOpen = new ObservableValue<boolean>(false);
    private membersSelected = new DropdownMultiSelection();
    private dropdownItems:Array<IListBoxItem<{}>> =[];

    constructor(props: {}) {
        super(props);
    }

    state: SprinPlannerState = {
        showMessage: false,
        selectedDuration: [],
        selectedMembers: [],
    };

    public componentDidMount() {
        SDK.init();
        this.loadAccounts();
    }

    public render(): JSX.Element {
        return (
            <Page className="sample-hub flex-grow">
                <Header title="Sprint Planner" />
                <div className="page-content">
                    <div className="sample-form-section flex-row flex-center">
                        <div className="flex-column">
                            <label htmlFor="work-item-type-picker">Sprint Duration in weeks:</label>
                            <Dropdown<string>
                                className="sample-work-item-type-picker"
                                items={[
                                    { id: "duration1", text: "1 week" },
                                    { id: "duration2", text: "2 weeks" },
                                    { id: "duration3", text: "3 weeks" },
                                    { id: "duration4", text: "4 weeks" },
                                ]}
                                onSelect={(event, item) => { this.state.selectedDuration.push(item.text || "")}}
                                selection={this.selection}
                            />
                        </div>
                    </div>
                    <div className="sample-form-section flex-row flex-center">
                        <div className="flex-column">
                            <label htmlFor="account-picker">Team members:</label>
                                <div style={{ margin: "8px" }}>
                                    <Observer selection={this.membersSelected}>
                                        {() => {
                                            return (
                                                <Dropdown
                                                    ariaLabel="Multiselect"
                                                    actions={[
                                                        {
                                                            className: "bolt-dropdown-action-right-button",
                                                            disabled: this.membersSelected.selectedCount === 0,
                                                            iconProps: { iconName: "Clear" },
                                                            text: "Clear",
                                                            onClick: () => {
                                                                this.membersSelected.clear();
                                                            }
                                                        }
                                                    ]}
                                                    className="example-dropdown"
                                                    items={this.dropdownItems}
                                                    selection={this.membersSelected}
                                                    placeholder="Select an Option"
                                                    showFilterBox={true}
                                                />
                                            );
                                        }}
                                    </Observer>
                                </div>
                            <Button className="sample-work-item-button" text="Preview Sprint" onClick={() => {
                                    this.prepareSelectedItems();
                                    console.log("printing selected members for dropdown+ ankit " + this.state.selectedMembers)
                                    this.setState(() => ({ showMessage: true }))
                                }} />
                            {this.state.showMessage && <SprintPreviewContent {...{status: this.state.showMessage, selectedMembers: this.state.selectedMembers, selectedDuration: this.state.selectedDuration ,callback: (status:boolean)=>{this.setState(() => ({ selectedDuration:[],showMessage: status,selectedMembers:[] }))}}}/>}
                        </div>
                    </div>
                </div>
            </Page>
        );
    }

    private prepareSelectedItems():void{
        console.log(" value of memberselected from planner  " +this.membersSelected.value);
        for(var i =0; i< this.membersSelected.value.length;i++){
           var temp = this.membersSelected.value[i]; 
           console.log(" value of temp from planner  " +temp);
           for(var j=temp.beginIndex;j<=temp.endIndex;j++) {
               var name = this.dropdownItems[j].text
               if(name != undefined)
               this.state.selectedMembers.push(name);
           }
        }
    }

    private async loadAccounts(): Promise<void> {

        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        let teamMemberNames: string[] = [];

        if (project)  {
            const client = await getClient(CoreRestClient);
            const teamMembers = await client.getTeamMembersWithExtendedProperties(project.name, "MSHackathon2020");
            teamMemberNames = teamMembers.map(t => t.identity.displayName);
            teamMemberNames.sort((a, b) => localeIgnoreCaseComparer(a, b));
        }

        this.teamMembers.change(0, ...teamMemberNames);
        for(var i =0; i< teamMemberNames.length;i++){
            console.log("Printing name from SprintPlanner +ankit "+teamMemberNames[i])
            var item = {
                id:teamMemberNames[i],
                text:teamMemberNames[i]
            }
            this.dropdownItems[i] = item;
        }
    }
}

showRootComponent(<SprintPlannerContent />);
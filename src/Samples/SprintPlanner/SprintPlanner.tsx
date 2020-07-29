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

import { CommonServiceIds, getClient, IProjectPageService } from "azure-devops-extension-api";
import { showRootComponent } from "../../Common";
import { CoreRestClient } from "azure-devops-extension-api/Core";
import SprintPreviewContent from "./SprintPreview"

const firstCheckbox = new ObservableValue<boolean>(false);
const secondCheckbox = new ObservableValue<boolean>(false);
const thirdCheckbox = new ObservableValue<boolean>(false);


type SprinPlannerState = {
    showMessage: boolean,
};

class SprintPlannerContent extends React.Component<{}, {}> {

    private workItemTypeValue = new ObservableValue("Bug");
    private selection = new ListSelection();
    private teamMembers = new ObservableArray<string>();
    private isDialogOpen = new ObservableValue<boolean>(false);

    constructor(props: {}) {
        super(props);
    }

    state: SprinPlannerState = {
        showMessage: false,
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
                                onSelect={(event, item) => { this.workItemTypeValue.value = item.data! }}
                                selection={this.selection}
                            />
                        </div>
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
                        <Button className="sample-work-item-button" text="Generate Sprint" onClick={() => {this.setState(() => ({ showMessage: true }))}} />
                        {this.state.showMessage && <SprintPreviewContent {...{status: this.state.showMessage, callback: (status:boolean)=>{this.setState(() => ({ showMessage: status }))}}}/>}
                    </div>
                </div>
            </Page>
        );
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
    }
}

showRootComponent(<SprintPlannerContent />);
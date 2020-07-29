import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import "./SprintPlanner.scss";

import { Button } from "azure-devops-ui/Button";
import { Header } from "azure-devops-ui/Header";
import { Page } from "azure-devops-ui/Page";


import { showRootComponent } from "../../Common";
import SprintPreviewContent from "./SprintPreview"

type QuoteState = {
    showMessage: boolean,
}

class SprintPlannerContent extends React.Component<{}, {}> {

    constructor(props: {}) {
        super(props);
    }

    state: QuoteState = {
        showMessage: false,
    };

    public componentDidMount() {
        SDK.init();
    }

    public render(): JSX.Element {
        return (
            <Page className="sample-hub flex-grow">
                <Header title="Sprint Planner" />
                <div className="page-content">
                    <div className="sample-form-section flex-row flex-center">
                        <Button className="sample-work-item-button" text="Generate Sprint" onClick={() => {this.setState(() => ({ showMessage: true }))}} />
                        {this.state.showMessage && <SprintPreviewContent {...{status: this.state.showMessage, callback: (status:boolean)=>{this.setState(() => ({ showMessage: status }))}}}/>}
                    </div>
                </div>
            </Page>
        );
    }
}
showRootComponent(<SprintPlannerContent />);
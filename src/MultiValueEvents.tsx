import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { getClient } from "TFS/WorkItemTracking/RestClient";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";

import { getSuggestedValues } from "./getSuggestedValues";
import { MultiValueControl } from "./MultiValueControl";

initializeIcons();
const HELP_URL = "https://github.com/Microsoft/vsts-extension-multivalue-control#azure-devops-services";
const DONE_STATES = ["Done", "Closed", "Resolved"];

export class MultiValueEvents {
    public readonly fieldName = VSS.getConfiguration().witInputs.FieldName;
    private readonly _container = document.getElementById("container") as HTMLElement;
    private _onRefreshed: () => void;
    /** Counter to avoid consuming own changed field events. */
    private _fired: number = 0;

    public async refresh(selected?: string[]): Promise<void> {
        let error = <></>;
        if (!selected) {
            if (this._fired) {
                this._fired--;
                if (this._fired !== 0) {
                    return;
                }
                error = await this._checkFieldType();
                if (!error) {
                    return;
                }
            }
            selected = await this._getSelected();
        }

        const dodError = await this._validateDoD(selected);
        if (dodError) {
            error = dodError;
        }

        ReactDOM.render(<MultiValueControl
            selected={selected}
            options={await getSuggestedValues()}
            onSelectionChanged={this._setSelected}
            error={error}
        />, this._container, () => {
            this._resize();
            if (this._onRefreshed) {
                this._onRefreshed();
            }
        });
    }
    private async _validateDoD(selected: string[]): Promise<JSX.Element | null> {
        try {
            const formService = await WorkItemFormService.getService();
            const state = await formService.getFieldValue("System.State") as string;
            if (DONE_STATES.indexOf(state) >= 0) {
                const options = await getSuggestedValues();
                const unchecked = options.filter((o) => selected.indexOf(o) < 0);
                if (unchecked.length > 0) {
                    return <div className="dod-validation-error">
                        {`All Definition of Done items must be completed before moving to "${state}". Missing: ${unchecked.join(", ")}`}
                    </div>;
                }
            }
        } catch (e) {
            // ignore validation errors to not break the control
        }
        return null;
    }

    private _resize = () => {
        VSS.resize(this._container.scrollWidth || 200, this._container.scrollHeight || 40);
    }
    private async _getSelected(): Promise<string[]> {
        const formService = await WorkItemFormService.getService();
        const value = await formService.getFieldValue(this.fieldName);
        if (typeof value !== "string") {
            return [];
        }
        return value.split(";").filter((v) => !!v).map(s => s.trim());
    }
    private _setSelected = async (values: string[]): Promise<void> => {
        const formService = await WorkItemFormService.getService();
        const fields = await formService.getFields();

        const currentField = fields.filter((f) => f.referenceName === this.fieldName)[0];

        if (!currentField) {
            console.warn(`Field ${this.fieldName} not found.`);
            return;
        }

        if (currentField.readOnly) {
            console.warn("Field is read only, cannot set value.");
            return;
        }
        this.refresh(values);
        this._fired++;
        await formService.setFieldValue(this.fieldName, values.join(";"));

        return new Promise<void>((resolve) => {
            this._onRefreshed = resolve;
        });
    }
    private async _checkFieldType(): Promise<JSX.Element> {
        const formService = await WorkItemFormService.getService();
        const inv = await formService.getInvalidFields();
        if (inv.length > 0 && inv.some((f) => f.referenceName === this.fieldName)) {
            const field = await getClient().getField(this.fieldName);
            if (field.isPicklist) {
                return <div>
                    {`Set the field ${field.name} to use suggested values rather than allowed values. `}
                    <a href={HELP_URL} target="_blank">{"See documentation"}</a>
                </div>;
            }
        }
        return <></>;
    }
}

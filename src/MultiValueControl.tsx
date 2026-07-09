import { Checkbox } from "office-ui-fabric-react/lib/components/Checkbox";
import {
  FocusZone,
  FocusZoneDirection,
} from "office-ui-fabric-react/lib/FocusZone";
import * as React from "react";

interface IMultiValueControlProps {
  selected?: string[];
  readOnly?: boolean;
  onSelectionChanged?: (selection: string[]) => Promise<void>;
  options: string[];
  error: JSX.Element;
  helpUrl?: string;
  helpUrlLabel?: string;
}

export class MultiValueControl extends React.Component<IMultiValueControlProps> {
  private _toggleOption = (option: string): void => {
    const selected = this.props.selected || [];
    const newSelected = selected.indexOf(option) >= 0
      ? selected.filter((o) => o !== option)
      : [...selected, option];
    if (this.props.onSelectionChanged) {
      this.props.onSelectionChanged(newSelected);
    }
  };

  public render() {
    const { selected = [], options, error, readOnly, helpUrl, helpUrlLabel } = this.props;
    const linkLabel = helpUrlLabel || "Definition of Done guidelines";

    return (
      <div className="multi-value-control dod-checklist" style={{ width: "100%", padding: "4px 0" }}>
        {helpUrl && (
          <div className="dod-help-link-container">
            <a href={helpUrl} target="_blank" rel="noopener noreferrer" className="dod-help-link">
              📖 {linkLabel}
            </a>
          </div>
        )}
        <FocusZone direction={FocusZoneDirection.vertical}>
          {options.map((o) => (
            <Checkbox
              key={o}
              className="text dod-item"
              checked={selected.indexOf(o) >= 0}
              disabled={readOnly}
              onChange={() => this._toggleOption(o)}
              label={o}
              title={o}
            />
          ))}
        </FocusZone>
        <div className="error">{error}</div>
      </div>
    );
  }
}

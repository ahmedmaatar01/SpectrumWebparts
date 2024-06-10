import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IDropdownOption } from "office-ui-fabric-react";
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle, PropertyFieldMultiSelect } from "@pnp/spfx-property-controls";

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'Ged365WebpartWebPartStrings';
import Ged365Webpart from './components/Ged365Webpart';
import { IGed365WebpartProps } from './components/IGed365WebpartProps';
import { SPOperations, SPListColumn } from '../Services/SPServices';

export interface IGed365WebpartWebPartProps {
  description: string;
  list_titles: IDropdownOption[];
  list_title: string;
  backgroundColor: string;
  textColor: string;
  selectedColumns: string[];
  columnOptions: IDropdownOption[]; // New property for storing column options
}

export default class Ged365WebpartWebPart extends BaseClientSideWebPart<IGed365WebpartWebPartProps> {

  private _spOperations: SPOperations;

  constructor() {
    super();
    this._spOperations = new SPOperations();
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IGed365WebpartProps> = React.createElement(
      Ged365Webpart,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        list_title: this.properties.list_title,
        backgroundColor: this.properties.backgroundColor,
        textColor: this.properties.textColor,
        selectedColumns: this.properties.selectedColumns
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;

      // Set default background color if not set
      if (!this.properties.backgroundColor) {
        this.properties.backgroundColor = '#3c3b5e';
      }
      if (!this.properties.textColor) {
        this.properties.textColor = '#ffffff';
      }
      if (!this.properties.selectedColumns || this.properties.selectedColumns.length === 0) {
        this.properties.selectedColumns = ['FileLeafRef'];
      }

      // Fetch all lists
      return this._spOperations.GetAllList(this.context)
        .then((result: IDropdownOption[]) => {
          this.properties.list_titles = result;

          // Fetch columns for the selected list title
          if (this.properties.list_title) {
            return this._fetchColumns(this.properties.list_title);
          }
        });
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('list_title', {
                  label: "select a title",
                  options: this.properties.list_titles,
                  selectedKey: this.properties.list_title,
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldMultiSelect('selectedColumns', {
                  key: 'selectedColumns',
                  label: 'Select columns to display',
                  options: this.properties.columnOptions || [], // Use dynamic column options
                  selectedKeys: this.properties.selectedColumns
                }),
                PropertyFieldColorPicker('backgroundColor', {
                  label: "Select background color",
                  selectedColor: this.properties.backgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'backgroundColorFieldId'
                }),
                PropertyFieldColorPicker('textColor', {
                  label: "Select text color",
                  selectedColor: this.properties.textColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: 'textColorFieldId'
                })
      
              ]
            }
          ]
        }
      ]
    };
  }

  private _fetchColumns(listTitle: string): Promise<void> {
    return this._spOperations.GetListColumns(this.context, listTitle)
      .then((columns: SPListColumn[]) => {
        const excludedColumns = ["Title", "_ExtendedDescription", "ContentType"];
        const columnOptions = columns
          .filter(column => !excludedColumns.includes(column.internalName))
          .map(column => ({
            key: column.internalName,
            text: column.title
          }));
        this.properties.columnOptions = columnOptions;
  
        this.context.propertyPane.refresh(); // Refresh the property pane to show new column options
      })
      .catch(error => {
        console.error('Error fetching columns:', error);
      });
  }
  

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'backgroundColor' && newValue !== oldValue) {
      this.properties.backgroundColor = newValue;
      this.render();
    } else if (propertyPath === 'list_title' && newValue !== oldValue) {
      this.properties.list_title = newValue;
      this._fetchColumns(newValue).then(() => {
        this.context.propertyPane.refresh();
        this.render();
      });
    } else if (propertyPath === 'textColor' && newValue !== oldValue) {
      this.properties.textColor = newValue;
      this.render();
    } else if (propertyPath === 'selectedColumns' && newValue !== oldValue) {
      this.properties.selectedColumns = newValue;
      this.render();
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}

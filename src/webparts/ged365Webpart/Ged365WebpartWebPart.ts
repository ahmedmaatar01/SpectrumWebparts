import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IDropdownOption } from "office-ui-fabric-react";

import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  
} from '@microsoft/sp-property-pane';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from "@pnp/spfx-property-controls";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'Ged365WebpartWebPartStrings';
import Ged365Webpart from './components/Ged365Webpart';
import { IGed365WebpartProps } from './components/IGed365WebpartProps';
import { SPOperations } from '../Services/SPServices';

export interface IGed365WebpartWebPartProps {
  description: string;
  list_titles : IDropdownOption[];
  list_title: string;
  displayType:string;
  backgroundColor:string;
}

export default class Ged365WebpartWebPart extends BaseClientSideWebPart<IGed365WebpartWebPartProps> {
  
  public componentWillMount() {
    this.getPropertyPaneConfiguration();
  }
  public _spOperations: SPOperations;
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
        displayType:this.properties.displayType
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.getPropertyPaneConfiguration();
    this.properties.displayType ="table";
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
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
    const {
      semanticColors
    } = currentTheme;

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
    this._spOperations
    .GetAllList(this.context)
    .then((result: IDropdownOption[]) => {
      this.properties.list_titles= result;
    })
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('list_title', {
                  label: "select a title",
                  options:this.properties.list_titles,
                }),
                PropertyPaneDropdown('displayType', {
                  label: "Select display type",
                  options: [
                    { key: 'table', text: 'Table' },
                    { key: 'grid', text: 'Grid' }
                  ],
                  selectedKey: 'table'
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
                  key: 'colorFieldId'
                  
                })
          ]
        }
      ]
    }
      ]
  };
}

}

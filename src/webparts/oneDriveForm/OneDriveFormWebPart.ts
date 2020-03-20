import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'OneDriveFormWebPartStrings';
import OneDriveForm from './components/OneDriveForm';
import { IOneDriveFormProps } from './components/IOneDriveFormProps';

export interface IOneDriveFormWebPartProps {
  description: string;
}

export default class OneDriveFormWebPart extends BaseClientSideWebPart<IOneDriveFormWebPartProps> {

  

  public render(): void {
    const element: React.ReactElement<IOneDriveFormProps > = React.createElement(
      OneDriveForm,
      {
        description: this.properties.description,
        webpartContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

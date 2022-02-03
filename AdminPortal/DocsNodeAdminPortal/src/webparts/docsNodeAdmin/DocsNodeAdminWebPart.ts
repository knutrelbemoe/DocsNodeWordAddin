import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DocsNodeAdminWebPartStrings';
import DocsNodeAdmin from './components/DocsNodeAdmin';
import { IDocsNodeAdminProps } from './components/IDocsNodeAdminProps';

export interface IDocsNodeAdminWebPartProps {
  description: string;
}

export default class DocsNodeAdminWebPart extends BaseClientSideWebPart<IDocsNodeAdminWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDocsNodeAdminProps > = React.createElement(
      DocsNodeAdmin,
      {
        description: this.properties.description,
        context: this.context
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

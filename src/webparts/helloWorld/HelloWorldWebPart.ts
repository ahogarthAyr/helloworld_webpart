import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
  EncodedAbsUrl: string;
  Description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.helloWorld }">
        <div id="spListContainer" />
    </div>`;
  
  this._renderListAsync();
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
                label: 'Description'
              }),
              PropertyPaneTextField('test', {
                label: 'Multi-line Text Field',
                multiline: true
              }),
              PropertyPaneCheckbox('test1', {
                text: 'Checkbox'
              }),
              PropertyPaneDropdown('test2', {
                label: 'Dropdown',
                options: [
                  { key: '1', text: 'One' },
                  { key: '2', text: 'Two' },
                  { key: '3', text: 'Three' },
                  { key: '4', text: 'Four' }
                ]}),
              PropertyPaneToggle('test3', {
                label: 'Toggle',
                onText: 'On',
                offText: 'Off'
              })
            ]
            }
          ]
        }
      ]
    };
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Site%20Pages')/items?$select=EncodedAbsUrl,Title,Description`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `       
              <div class="${styles.column}">
                  <a class="${styles.title} "href="${item.EncodedAbsUrl}">${item.Title}</a>
                  <div class="${styles.description}" >${item.Description}</div>
              </div>  
    `;
    });

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    
    if (Environment.type == EnvironmentType.SharePoint ||
             Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

}

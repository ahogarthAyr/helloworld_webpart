import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IHelloWorldWebPartProps {
  DropDownProp: string;
}

export interface ISPListItems {
  value: ISPListItem[];
}

export interface ISPListItem {
  Title: string;
  Id: string;
  EncodedAbsUrl: string;
  Description: string;
}

export interface spList{  
  Title:string;  
  Id: string;  
  }  
  export interface spLists{  
    value: spList[];  
  }  

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private dropDownOptions: IPropertyPaneDropdownOption[] = [];  
  // public constructor(context: WebPartContext) {  
  //   // super(context);  
  // }  


  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.helloWorld }">
        <div id="spListContainer" />
    </div>`;
    console.log("Render");  
  
    this.LoadData();  
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }  

  protected onPropertyPaneConfigurationStart(): void {  
    // Stops execution, if the list values already exists  
   if(this.dropDownOptions.length>0) return;  
 
   this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'DropDownProp');

   // Calls function to append the list names to dropdown  
   this.GetLists()  
 
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
 
 }  

 private GetLists():void{  
  // REST API to pull the list names  
  let listresturl: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=Id,title&$filter=Hidden ne true`;  

  this.LoadLists(listresturl).then((response)=>{  
    // Render the data in the web part  
    this.LoadDropDownValues(response.value);  
  });  
}  

private LoadLists(listresturl:string): Promise<spLists>{  


  return this.context.spHttpClient.get(listresturl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
    return response.json();  
  });  
}  

private LoadDropDownValues(lists: spList[]): void{  
  lists.forEach((list:spList)=>{  
    // Loads the drop down values  
    this.dropDownOptions.push({"key":list.Title,"text":list.Title});
    console.log(list.Title);
  });  
};



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName:"Lists",
              groupFields: [  
                PropertyPaneDropdown('DropDownProp',{  
                  label: "Select List To Display on the page",  
                  options: this.dropDownOptions,  
                  disabled: false,
                  // selectedKey: this.properties.DropDownProp  
                  
                })
            ]
            }
            
          ]
        }
      ]
    };
  }

  private GetListData(): Promise<ISPListItems> {
    let url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${this.properties.DropDownProp}')/items?$select=EncodedAbsUrl,Title,Description`

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  private RenderListData(items: ISPListItem[]): void {
    let html: string = '';
    items.forEach((item: ISPListItem) => {
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

  private LoadData(): void {

    // if(this.properties.DropDownProp != undefined){  
    if (Environment.type == EnvironmentType.SharePoint ||
             Environment.type == EnvironmentType.ClassicSharePoint) {
              this.GetListData().then((response)=>{  
                // Render the data in the web part  
                this.RenderListData(response.value);  
      
              });  
    }
    }
  // }

}

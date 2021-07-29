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
import { getItemStyles } from '@fluentui/react/lib/components/ContextualMenu/ContextualMenu.classNames';

export interface IHelloWorldWebPartProps {
  DropDownProp: string;
}

export interface ISPListItems {
  value: ISPListItem[];
}

export interface ISPListItem {
  Title: string;
  Id: number;
  EncodedAbsUrl: string;
  Description: string;
  PromotedState: number;
  ShowInListView: Boolean;
  WelcomePage: string;
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
  private listsDropdownDisabled: boolean = true;


  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.helloWorld }">
        <div id="spListContainer" />
        <div class="${styles.icon}"><i class="ms-Icon ms-Icon--CustomList" aria-hidden="false"></i><br/></div>
        <div class="${styles.ddSelect}">
        Select a list to add to this page.
        </div>
    </div>`; 
    // this.LoadData();
    this.LoadMostViewed();  
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }  


  protected onPropertyPaneConfigurationStart(): void {  
    // Stops execution, if the list values already exists  
   this.listsDropdownDisabled = !this.dropDownOptions;
   
    if(this.dropDownOptions.length>0){ 
     console.log('yes');  
     return;
   }

   this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'DropDownProp');


   // Calls function to append the list names to dropdown  

     this.GetLists();  
 
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        // this.context.statusRenderer.clearLoadingIndicator(this.domElement);
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
  });  
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
              groupName:"",
              groupFields: [  
                PropertyPaneDropdown('DropDownProp',{  
                  label: "Select List to Display on the page",  
                  options: this.dropDownOptions,  
                  disabled: this.listsDropdownDisabled,
                  selectedKey: this.properties.DropDownProp  
                  
                })
            ]
            }
            
          ]
        }
      ]
    };
  }

  // private GetListData(): Promise<ISPListItems> {
  //   let url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${this.properties.DropDownProp}')/items?$select=Id,EncodedAbsUrl,Title,Description,PromotedState,ShowInListView`;

  //   return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) => {
  //       return response.json();
  //     });
  // }
  // private RenderListData(items: ISPListItem[]): void {

  //   let html: string = '';
  //   items.forEach((item: ISPListItem) => {

  //     if(item.Id != 1 && item.PromotedState == 0 && item.ShowInListView == true){
  
  //     html += `       
  //             <div class="${styles.column}">
  //                 <a class="${styles.title} "href="${item.EncodedAbsUrl}">${item.Title}</a>
  //                 <div class="${styles.description}" >${item.Description}</div>
  //             </div>  
  //   `;  
  //     }
  //   });
    
  //   const listContainer: Element = this.domElement.querySelector('#spListContainer');
  //   listContainer.innerHTML = html;
  // }

     
  // private LoadData(): void {

  //   if(this.properties.DropDownProp != undefined){  
  //   if (Environment.type == EnvironmentType.SharePoint ||
  //            Environment.type == EnvironmentType.ClassicSharePoint) {
  //             this.GetListData().then((response)=>{  
  //               // Render the data in the web part  
  //               this.RenderListData(response.value);  
  //               this.context.propertyPane.refresh();

  //             });  
  //   }
  //   }
  // }

  
  private GetMostViewed(): Promise<any> {

    // query site pages for ViewsLifetime, sort descending and select properties to filter results

    let url = this.context.pageContext.web.absoluteUrl + 
    `/_api/search/query?querytext=%27path:https://ayrsandbox.sharepoint.com/SitePages%27&rowlimit=10&sortlist=%27ViewsLifetime:descending%27&selectproperties=%27DefaultEncodingUrl,%20Title,%20Description,%20promotedstate,%20ShowInListView%27`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();        
      });
    }
  


  private RenderMostViewed(items: any): any {

    
    let html: string = '';

    for(var i=0;i<items.length;i++){  

      if (items[i].Cells[5]["Value"] == 0 && items[i].Cells[6]["Value"] == 'true'){

     html += 
     `       
              <div class="${styles.column}">
                  <a class="${styles.title} "href="${items[i].Cells[2]["Value"]}">${items[i].Cells[3]["Value"]}</a>
                  <div class="${styles.description}" >${items[i].Cells[4]["Value"]}</div>
              </div>  

    `;  
      }
    };
    
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }


  private LoadMostViewed(): void {

    this.GetMostViewed().then((data)=>{
     
      let listItems = data.PrimaryQueryResult.RelevantResults.Table.Rows

      console.log(listItems)

      if(this.properties.DropDownProp == 'Site Pages'){

        this.RenderMostViewed(listItems);  
        this.context.propertyPane.refresh();
      }
    })
  }

}

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

export interface IHelloWorldWebPartProps {
  DropDownProp: string;
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
    this.LoadViews();  
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


private GetPageUrls():Promise<any>{
  
  let docLibrary = 'IT%20Support%20page%20links'
  let pageUrl: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${docLibrary}')/items?$select=Title,OData__ExtendedDescription,OData__ShortcutUrl`;

  return this.context.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
    return response.json();  
  }); 

}

private RenderPageUrls(items: any): any {
    
  let html: string = '';

  for(var i=0;i<items.length;i++){  

   html += 
   `       
            <div class="${styles.column}">
                <a class="${styles.title} "href="${items[i].OData__ShortcutUrl.Url}">${items[i].Title}</a>
                <div class="${styles.description}" >${items[i].OData__ExtendedDescription}</div>
            </div>  
  `;  
  };
  
  const listContainer: Element = this.domElement.querySelector('#spListContainer');
  listContainer.innerHTML = html;
}


private LoadPageUrls(): void {

  this.GetPageUrls().then((data)=>{

    // console.log(data.value)
    this.RenderPageUrls(data.value)

  })
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

  
  private GetMostViewed(): Promise<any> {

    // query site pages for ViewsLifetime, sort descending and select properties to filter results

    let absUrl = this.context.pageContext.site.absoluteUrl + '/SitePages'

    let url = this.context.pageContext.web.absoluteUrl + 
    `/_api/search/query?querytext=%27path:${absUrl} ShowInListView:yes%27&rowlimit=30&sortlist=%27ViewsLifetime:descending%27&selectproperties=%27DefaultEncodingUrl,%20Title,%20Description,%20promotedstate,ShowInListView%27`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();        
      });
    }
  

  private RenderMostViewed(items: any): any {
    
    let html: string = '';

    for(var i=0;i<items.length;i++){  

      if (items[i].Cells[5]["Value"] == 0){

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


  private LoadViews(): void {

    this.GetMostViewed().then((data)=>{
     
      let listItems = data.PrimaryQueryResult.RelevantResults.Table.Rows

      console.log(listItems)

      if(this.properties.DropDownProp == 'Site Pages'){

        this.RenderMostViewed(listItems);  
        this.context.propertyPane.refresh();
      }
      else if(this.properties.DropDownProp == 'IT Support page links') {

        this.LoadPageUrls()
        // console.log(this.LoadPageUrls())
      }
    })
  }

}

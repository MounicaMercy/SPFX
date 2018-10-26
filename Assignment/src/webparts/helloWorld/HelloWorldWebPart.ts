import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import IHelloWorldWebPartProps from './File';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
              <div>
              
              <div class="${ styles.container }">
              <div class="${ styles.row }"style="background-color:${this.properties.color}">
              <div class="${ styles.column }">
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
                Select Category:<select id="categories"></select>
              <table id="display">
               </table>
              </div>
              </div>
              </div>
              </div>
           
          `;
          this.getListsInfo();
          this.AddEventListeners();  //creating a event handler
  }
  private AddEventListeners() : void{
    document.getElementById('categories').addEventListener('change',()=>this.GetProducts()); 
  }
  public GetProducts()
  {

   let html: string = '';
   var selectedvalue=(<HTMLSelectElement>this.domElement.querySelector('#categories')).value;
     if (Environment.type === EnvironmentType.Local) {
     this.domElement.querySelector('#categories').innerHTML = "Sorry this does not work in local workbench";
   } else {
   this.context.spHttpClient.get
   (
     this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Products')/items?$filter=Category/Title eq'${selectedvalue}'`, 
     SPHttpClient.configurations.v1)
     .then((response: SPHttpClientResponse) => {
       response.json().then((listsObjects: any) => {
         listsObjects.value.forEach(listObject => {
           html +=`
           <tr>
           <td>
           ${listObject.Title}
           </td>
           </tr>`;
         });
         this.domElement.querySelector('#display').innerHTML = html;
       });
     });        
   }
  }
  private getListsInfo() {
    let html: string = '';
    if (Environment.type === EnvironmentType.Local) {
      this.domElement.querySelector('#categories').innerHTML = "Sorry this does not work in local workbench";
    } else {
    this.context.spHttpClient.get
    (
      this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('Category')/items?$select=Title,ID`, 
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((listsObjects: any) => {
          listsObjects.value.forEach(listObject => {
            html +=`<option value="${listObject.Title}">${listObject.Title}</option>`;
          });
          this.domElement.querySelector('#categories').innerHTML = html;
        });
      });        
    }
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
                }),
                PropertyPaneDropdown('color',{
                  label:"Select Color",
                  options:[
                    {key:"red",text:"Red"} ,
                    {key:"blue",text:"Blue"},
                    {key:"white",text:"White"},
                    {key:"green",text:"Green"}
                ] 
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

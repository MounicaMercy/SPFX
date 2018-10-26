import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as $ from 'jquery';

export interface IHelloWorldWebPartProps {
  description: string;
  
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              
            Select Category:<select id="categories"></select>
            <table id="display">
             </table>
            </div>
          </div>
        </div>
      </div>`;
      this.getListsInfo();
      this.AddEventListeners();
  }
  private AddEventListeners() : void{
    document.getElementById('categories').addEventListener('change',()=>this.GetProducts()); 
  }
  public GetProducts()
  {
  var selectedvalue=$("#categories").val();
  // var selectedvalue=(<HTMLSelectElement>this.domElement.querySelector('#categories')).value;
     if (Environment.type === EnvironmentType.Local) {
     this.domElement.querySelector('#categories').innerHTML = "Sorry this does not work in local workbench";
   } else {
    var call = $.ajax({
      url: this.context.pageContext.web.absoluteUrl +`/_api/web/Lists/getByTitle('Products')/items?$filter=Category/Title eq'${selectedvalue}'`,
      type: "GET",
        dataType: "json",
        headers: {
            Accept: "application/json;odata=verbose"
        }
      
    });
  call.done(function (data, textStatus, jqXHR) {
    var display = $("#display");

   
    $.each(data.d.results, function (index, value) {
      display.append(value.Title);
      display.append("<br/>");
        
    });
  });
  call.fail(function (jqXHR, textStatus, errorThrown) {
    var response = JSON.parse(jqXHR.responseText);
    var message = response ? response.error.message.value : textStatus;
    alert("Call failed. Error: " + message);
  });
   }
  }
  private getListsInfo() {
    //var url:this.context.pageContext.web.absoluteUrl;
 
    if (Environment.type === EnvironmentType.Local) {
      this.domElement.querySelector('#categories').innerHTML = "Sorry this does not work in local workbench";
    } 
    else 
    {
      var call = $.ajax({
        url: this.context.pageContext.web.absoluteUrl+`/_api/web/Lists/getByTitle('Category')/Items?$select=Title,ID`,
        type: "GET",
          dataType: "json",
          headers: {
              Accept: "application/json;odata=verbose"
          }
        
      });
    call.done(function (data, textStatus, jqXHR) {
      var CategoryList = $("#categories");
  
      $.each(data.d.results, function (index, value) {
        CategoryList.append("<option value="+value.Title+">"+value.Title+"</option>");
        
      });
    });
    call.fail(function (jqXHR, textStatus, errorThrown) {
      var response = JSON.parse(jqXHR.responseText);
      var message = response ? response.error.message.value : textStatus;
      alert("Call failed. Error: " + message);
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

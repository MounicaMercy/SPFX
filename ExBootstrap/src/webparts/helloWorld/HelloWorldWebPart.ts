import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import{SPComponentLoader}from '@microsoft/sp-loader';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import * as $ from 'jquery';
require('bootstrap');
export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    let url="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(url);
    this.domElement.innerHTML = `
      <div class="row">
        <div class="container">
          <div id="myCarousel" class="carousel slide" data-ride="carousel">

              <!-- Indicators -->
              <ol class="carousel-indicators">
                <li data-target="#myCarousel" data-slide-to="0" class="active"></li>
                <li data-target="#myCarousel" data-slide-to="1"></li>
              </ol>
             <!-- Column Dividing -->

              <div class="col-sm-8">
           
                <!-- Wrapper for slides -->

                <div class="carousel-inner" id="slider">
                
                  <!-- Second Column Divided -->

                  <div class="col-sm-4" style="background-color:lavenderblush;"> </div>

                  

                   <!-- Trigger the modal with a button 
                    <button type="button" class="btn btn-info btn-lg" data-toggle="modal" data-target="#myModal">Open Modal</button>

                    <!-- Modal -->

                    <div class="modal fade" id="myModal" role="dialog">
                      <div class="modal-dialog">
                        
                          <!-- Modal content-->
        
                          <div class="modal-content">
                            <div class="modal-header">
                              <button type="button" class="close" data-dismiss="modal">&times;</button>
                              <h4 class="modal-title">Modal Header</h4>
                            </div>
                            <div class="modal-body">
                              <p>Some text in the modal.</p>
                            </div>
                            <div class="modal-footer">
                              <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                            </div>
                          </div>-->  
                            
                      </div>
                    </div> 
                    
                </div>
              </div>
            </div>
          </div>
      </div>`
      ;
      this.getListsInfo();
      //this.AddEventListeners();
      $(document).ready(function () {

      });
  }
  private AddEventListeners() : void{
    document.getElementById('url').addEventListener('change',()=>this.GetProducts()); 
  }
  public GetProducts()
    {
    var selectedvalue=$("#categories").val();
    if (Environment.type === EnvironmentType.Local) 
    {
      this.domElement.querySelector('#categories').innerHTML = "Sorry this does not work in local workbench";
    } 
    else 
    {
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
   if (Environment.type === EnvironmentType.Local) 
    {
      this.domElement.querySelector('#slider').innerHTML = "Sorry this does not work in local workbench";
    } 
    else 
    {
      var call = $.ajax({
        url: this.context.pageContext.web.absoluteUrl+`/_api/web/Lists/getByTitle('Managers Speaks')/Items?$select=ImageUrl,ID,Subject,Description`,
        type:"GET",
          dataType: "json",
          headers: {
              Accept: "application/json;odata=verbose"
          }
      });
      call.done(function (data, textStatus, jqXHR) {
        var slidervar = $("#slider");
       var slidecounter=0;
        $.each(data.d.results, function (index, value) {
          if(slidecounter == 0)
          {
            slidervar.append(`
            <div class="item active">
             <img src="${value.ImageUrl}" alt="Los Angeles" style="width:100%;">
            </div>
            <div class ="carousel-caption"><p>"${value.Subject}"</p>
            </div>
            <!-- Controls -->
                  <a class="left carousel-control" href="#myCarousel" data-slide="prev">
                    <span class="glyphicon glyphicon-chevron-left"></span>
                    <span class="sr-only">Previous</span>
                  </a>
                  <a class="right carousel-control" href="#myCarousel" data-slide="next">
                    <span class="glyphicon glyphicon-chevron-right"></span>
                    <span class="sr-only">Next</span>
                  </a>
           `);
            slidecounter++;
          }
          else
          {
            slidervar.append(`
            <div class="item">
             <img src="${value.ImageUrl}" alt="Los Angeles" style="width:100%;">
            </div>
            <div class ="carousel-caption"><p>"${value.Subject}"</p>
            </div>
            <!-- Controls -->
                  <a class="left carousel-control" href="#myCarousel" data-slide="prev">
                    <span class="glyphicon glyphicon-chevron-left"></span>
                    <span class="sr-only">Previous</span>
                  </a>
                  <a class="right carousel-control" href="#myCarousel" data-slide="next">
                    <span class="glyphicon glyphicon-chevron-right"></span>
                    <span class="sr-only">Next</span>
                  </a>
            `);
          }
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

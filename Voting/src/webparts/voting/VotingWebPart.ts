import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

//Importing Packages
import{SPComponentLoader} from '@microsoft/sp-loader';
import styles from './VotingWebPart.module.scss';
import * as strings from 'VotingWebPartStrings';
import * as $ from 'jquery';
import pnp from 'sp-pnp-js';
import Chart from 'chart.js';
require('bootstrap');

export interface IVotingWebPartProps {
  description: string;
}

// Global Variables
var Locationid;
var CurrentUser;
var IsVoted:boolean;
var CurrentUserID;

export default class VotingWebPart extends BaseClientSideWebPart<IVotingWebPartProps> {
  public render(): void {
    let url="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(url);
    CurrentUser = this.context.pageContext.user.displayName;
    this.CheckingUser();  //Checking the user as entered the page
    this.domElement.innerHTML = `
      <div class="${ styles.voting }">
        <div class="${ styles.container }">
            <div class="${ styles.column }">
              <div id="Displayusername"></div>
              <div id="display"></div>
              </br>
              <button type="button" id="saveid">SAVE</button>
              </br>
              <div id="PieChart">
              <canvas id="pie-chart" width="800" height="450"></canvas> 
              </div>
            </div>
        </div>
      </div>`;
      this.GetLocationButton(); //Getting and Creating the Buttons
      $(document).ready(function()
      {
       alert("documet alert");
        $(document).on("click",".btn",function() //On Click-Location Buttons
        {
          //Getting the slected button id
          Locationid=$(this).attr('id');
          //Bootstrap for getting highlighted
          $(".btn").removeClass('active').addClass('disabled');
          $('#'+Locationid).removeAttr('class');
          $('#'+Locationid).addClass('active btn btn-success');
        });
        $(document).on("click","#saveid",function() //On Click-Save Button
        {
          alert(Locationid);
          if(IsVoted==true)
          {
            Updated(); //If user already voted then update!
          }
          else if (IsVoted ==false)
          {
            SaveVote(); //If a new user then inserting a new item
          }
        });
      });
      function Updated()
      {
        if (Environment.type === EnvironmentType.Local) //Checking Environment
          {
            this.domElement.querySelector('#saveid').innerHTML = "Sorry this does not work in local workbench";
          }
        else
          {   //Updating the list by current user id
            pnp.sp.web.lists.getByTitle("Mounica_Votes").items.getById(CurrentUserID).update({
              Title: Locationid
          })
          }
          alert("Updating Vote..");
      }
      function SaveVote()
      {
        if (Environment.type === EnvironmentType.Local)  //Checking Environment
          {
            this.domElement.querySelector('#saveid').innerHTML = "Sorry this does not work in local workbench";
          }
        else
          {    //Inserting an item to the list
            pnp.sp.web.lists.getByTitle("Mounica_Votes").items.add({
            Title: Locationid,
            User:CurrentUser
            });
          }
          alert("Saving vote..");
      }
  }
  private CheckingUser()
      {
        if(Environment.type === EnvironmentType.Local)  //Checking Environment
          {
            this.domElement.querySelector('#Displayusername').innerHTML = "Sorry this does not work in local workbench";
          }
        else
          {
            alert("Checking user entered");
            var call = $.ajax({
              url: this.context.pageContext.web.absoluteUrl+`/_api/web/Lists/getByTitle('Mounica_Votes')/Items?$select=Title,User,ID&$filter=User eq '${CurrentUser}'`,
              type:"GET",
                dataType: "json",
                headers: {
                    Accept: "application/json;odata=verbose"
                }
            });
            call.done(function (data, textStatus, jqXHR) {
              var GetUser = $("#Displayusername");
              //setting it to false so as to insert a new item
              IsVoted=false;
              //Checking thye user and highlighting the already selected location
              $.each(data.d.results, function (index, value) {
                //displaying the selected location id/number
                GetUser.append(`"You already voted to" ${value.Title}`);
                alert("disable the buttons");
                $(".btn btn-success").removeClass('active').addClass('disabled');
                $('#'+value.Title).removeAttr('class');
                $('#'+value.Title).addClass('active btn btn-success');
                //Saving the id
                CurrentUserID=`${value.ID}`;
                alert("Current userid is "+CurrentUserID);
                //Setting it to true so as to update
                IsVoted=true;
              });
            });
            call.fail(function (jqXHR, textStatus, errorThrown) {
              var response = JSON.parse(jqXHR.responseText);
              var message = response ? response.error.message.value : textStatus;
              alert("Call failed. Error: " + message);
            });
          }
  }
  private GetLocationButton()
  {
    if (Environment.type === EnvironmentType.Local)   //Checking Environment
    {
      this.domElement.querySelector('#display').innerHTML = "Sorry this does not work in local workbench";
    }
    else
    {
      var call = $.ajax({
        url: this.context.pageContext.web.absoluteUrl+`/_api/web/Lists/getByTitle('Mounica_Location')/Items?$select=Location,ID`,
        type:"GET",
          dataType: "json",
          headers: {
              Accept: "application/json;odata=verbose"
          }
      });
      call.done(function (data, textStatus, jqXHR) {
        var location = $("#display");
        $.each(data.d.results, function (index, value) {
        location.append(`<button type="button" class="btn btn-success" id="${value.ID}">${value.Location}</button>&nbsp`);
        });
      });
      call.fail(function (jqXHR, textStatus, errorThrown) {
        var response = JSON.parse(jqXHR.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
      });
    }
    this.DrawPieChart();
  }
  private DrawPieChart()
  {
    new Chart(document.getElementById("pie-chart"), 
    {
    type: 'pie',
    data: {
    labels:['Hyderabad','Banglore','Goa'],
    datasets: 
    [
      {
      label: "Votes submitted",
      backgroundColor: ["#3e95cd", "#8e5ea2","#3cba9f","#e8c3b9","#c45850"],
      data: [25,14,23]
      }
    ]
    },
    options: 
    {
      title: 
      {
        display: true,
      }
    }
  });
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

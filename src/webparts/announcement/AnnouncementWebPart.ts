import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AnnouncementWebPartStrings';
import Announcement from './components/Announcement';
import { IAnnouncementProps } from './components/IAnnouncementProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from "jquery";

export interface IAnnouncementWebPartProps {
  description: string;
}

export default class AnnouncementWebPart extends BaseClientSideWebPart<IAnnouncementWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IAnnouncementProps> = React.createElement(
      Announcement,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
SPComponentLoader.loadCss('https://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css');
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js');
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.9.0/css/bootstrap-datepicker.min.css');
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.9.0/js/bootstrap-datepicker.min.js');
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js');
SPComponentLoader.loadScript('https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css');
SPComponentLoader.loadScript('https://cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js');
this.domElement.innerHTML = ` <label style="text-align: center;"><b><h3>Announcement View</h3></b></label>
<hr style="border: 3px solid black; width: 100%; margin: 10px 0;">
<div class="row mt-4 mb-5">
<button type="button" id="btnSubmit" style="width:180px;" class="btn btn-primary">Submit Annoucement</button>
</div>
<div class="row mt-4 mb-5">
<div class="col-sm-12">

    <table id='tblannounce' class="table table-bordered">
    <thead>
    <th>Title</th>
    <th>Description</th>
    <th>Date of posting</th>
    </thead>
    <tbody>
    </tbody>

    </table>
</div>
</div>

`;
const urlofsite =this.context.pageContext.web.absoluteUrl;
 $(() =>{
  loadDatatable();
     });

     const loadDatatable=()=>
      {
        let listname="Announcement"
        const endPointUrl = `${urlofsite}/_api/web/lists/getbytitle('${listname}')/items?$select=Title,Description,dateofannounce&$orderby=dateofannounce desc `;
        $.ajax({
          url: endPointUrl,
          type: 'GET',
          dataType: 'json',
          success: function (data) {
            const tableBody = $('#tblannounce tbody');
                   tableBody.empty();
           if(data.value.length>0)
            {
              data.value.forEach((item:any) => {
                const row = `<tr>
                <td>${item.Title}</td>
                <td>${item.Description}</td>
                <td>${item.dateofannounce}</td>
                </tr>`;
                tableBody.append(row);
              
              })
             
            }
          },
          error: function (error) {
              console.error('Error fetching status data:', error);
          }
      });
      }
  }

  

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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

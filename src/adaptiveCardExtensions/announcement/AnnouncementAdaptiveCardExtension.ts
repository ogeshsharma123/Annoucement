import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { AnnouncementPropertyPane } from './AnnouncementPropertyPane';
import { QuickView } from './quickView/QuickView';
import {setup as pnpSetup} from "@pnp/common";



export interface IAnnouncementAdaptiveCardExtensionProps {
  title: string;
}


export interface IListItems {
  Title: string;
  Description: string;
  dateofannounce: Date;
  id: number;
}
export interface IAnnouncementAdaptiveCardExtensionState {
  issueList: IListItems[];
  index:number;
}

const CARD_VIEW_REGISTRY_ID: string = 'Announcement_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Announcement_QUICK_VIEW';


export default class AnnouncementAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAnnouncementAdaptiveCardExtensionProps,
  IAnnouncementAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AnnouncementPropertyPane;

  public onInit(): Promise<void> {
    
    return super.onInit().then(() => {
      // Set up PnP in a way that ensures `this.context` is defined
      pnpSetup({
        spfxContext: this.context,
      });
    });
    this.GetListItems().then((announceListResponse) => {
      let announceList: IListItems[];
      announceList= [];
      announceListResponse.map((item:any) => {
        announceList.push({
            Title: item.Title,
            Description: item.Description,
            dateofannounce: item.dateofannounce,
              id: item.ID
          });
      });
      
  });
    

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }
  protected async GetListItems(): Promise < any > {
    const urlofsite =this.context.pageContext.web.absoluteUrl;
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
          return data;
         
        }
      },
      error: function (error) {
          console.error('Error fetching status data:', error);
      }
  });
}


  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Announcement-property-pane'*/
      './AnnouncementPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.AnnouncementPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}

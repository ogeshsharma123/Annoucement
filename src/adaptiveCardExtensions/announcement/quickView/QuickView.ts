import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
//import * as strings from 'AnnouncementAdaptiveCardExtensionStrings';
import {
  IAnnouncementAdaptiveCardExtensionProps,
  IAnnouncementAdaptiveCardExtensionState
} from '../AnnouncementAdaptiveCardExtension';

export interface IListItems {
  Title: string;
  Description: string;
  dateofannounce: Date;
  
}
export interface IQuickViewData {
  announceList:IListItems[]
}

export class QuickView extends BaseAdaptiveCardQuickView<
  IAnnouncementAdaptiveCardExtensionProps,
  IAnnouncementAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      announceList:this.state.issueList
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}

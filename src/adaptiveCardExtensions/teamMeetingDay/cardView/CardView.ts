import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TeamMeetingDayAdaptiveCardExtensionStrings';
import { ITeamMeetingDayAdaptiveCardExtensionProps, 
         ITeamMeetingDayAdaptiveCardExtensionState, 
         VIEW_QUICK_VIEW_REGISTRY_ID, EDIT_QUICK_VIEW_REGISTRY_ID } from '../TeamMeetingDayAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<ITeamMeetingDayAdaptiveCardExtensionProps, ITeamMeetingDayAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: VIEW_QUICK_VIEW_REGISTRY_ID
          }
        }
      },
      {
        title: strings.OpenEditBtn,
        action: {
          type: 'QuickView',
          parameters: {
            view: EDIT_QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IPrimaryTextCardParameters {

    const primaryText = this.state.user.team?.length>0 ?this.state.popularDay:strings.NoColleague;
    const desc = this.state.user.team?.length>0?  `${this.state.pupularDayCount} ${strings.PeoplesGoing}`;

    return {
      primaryText: primaryText ,
      description: desc
    };
  }

  // public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
  //   return {
  //     type: 'ExternalLink',
  //     parameters: {
  //       target: 'https://www.bing.com'
  //     }
  //   };
  // }
}

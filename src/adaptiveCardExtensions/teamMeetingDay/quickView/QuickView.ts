import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TeamMeetingDayAdaptiveCardExtensionStrings';
import { ITeamMeetingDayAdaptiveCardExtensionProps, ITeamMeetingDayAdaptiveCardExtensionState, IUser } from '../TeamMeetingDayAdaptiveCardExtension';

export interface IQuickViewData {
  title:string;
  peoples:IUser[];
}

export class QuickView extends BaseAdaptiveCardView<
  ITeamMeetingDayAdaptiveCardExtensionProps,
  ITeamMeetingDayAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {    
    console.log(typeof this.state.user.team[0].schedule.monday)

    return {
      title:strings.YourTeam,
      peoples:this.state.user?.team
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}

/*
"type": "ColumnSet",
          "columns": [
            {
              "type":"Column",
              "items":[
                {    
                  "type":"TextBlock",                   
                  "text": "${displayName}"
                }
              ]
            },
            {
              "type":"Column",
              "items":[
                {             
                  "type":"TextBlock",                   
                  "text": "${displayName}"
                }
              ]
            }
         ]



         {
  "schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.2",
  "body": [
    {
      "type": "TextBlock",
      "weight": "Bolder",
      "text": "${title}"
    },
    {
      "type": "FactSet",
      "$data":"${peoples}",
      "facts": [
        {                       
            "value": "${displayName}"
        },
        {
          "title": "Monday",
          "value": "ON"        
        }
      ],
      "separator": true 
    }
  ]
}


         */
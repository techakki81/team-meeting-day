import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension,CardSize } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { EditView } from './quickView/EditView';
import { TeamMeetingDayPropertyPane } from './TeamMeetingDayPropertyPane';
import { getUserTeamData } from '../../utilities/getUserData';
import { weekNumber } from '../../utilities/weekNumber';
import { DisplayMode, Environment } from '@microsoft/sp-core-library';


import { setup as pnpSetup } from "@pnp/common";
import { ISharePointQueryable, sp } from "@pnp/sp";
import "@pnp/common";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items"

import {ICamlQuery, IList}  from "@pnp/sp/lists";
import { IItem, IItemAddResult } from "@pnp/sp/items";
import { Constants } from '../../utilities/constants';


export interface ITeamMeetingDayAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
}

export interface ITeamMeetingDayAdaptiveCardExtensionState { 
  popularDay?:string;
  pupularDayCount?:number;
  user?:IUser;
  weekNumber:number;
  spList:IList;
  spListItem?:any
}
export interface ISchedule {
  monday?:string;
  tuesday?:string;
  wednesday?:string;
  thursday?:string;
  friday?:string;
}

export interface IUser{
  displayName:string
  email:string;
  schedule?:ISchedule;  
  team?:IUser[];
}


export const CARD_VIEW_REGISTRY_ID: string = 'TeamMeetingDay_CARD_VIEW';
export const VIEW_QUICK_VIEW_REGISTRY_ID: string = 'View_TeamMeetingDay_QUICK_VIEW';
export const EDIT_QUICK_VIEW_REGISTRY_ID: string = 'Edit_TeamMeetingDay_QUICK_VIEW';

 


export default class TeamMeetingDayAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITeamMeetingDayAdaptiveCardExtensionProps,
  ITeamMeetingDayAdaptiveCardExtensionState
> {

  private _deferredPropertyPane: TeamMeetingDayPropertyPane | undefined;

  public onInit(): Promise<void> {

   
    pnpSetup({
      spfxContext: this.context
    }); 

    
    const spLst = sp.web.getList(Constants.ListUrl)

    this.state = {      
      weekNumber : weekNumber( new Date()),
      spList: spLst
    };


    getUserTeamData(this.context,spLst).then( (val:IUser[]) =>{

      // return if there is no team
      if( !val || val.length<=0 )
        return

      // loop over the team schedule and get the most popular day and count 
      let dayCount ={
        Monday:0,
        Tuesday:0,
        Wednesday:0,
        Thursday:0,
        Friday:0        
      }

      val.forEach(  (usr:IUser) => {
        usr.schedule?.monday==="true" && dayCount.Monday++
        usr.schedule?.tuesday==="true" && dayCount.Tuesday++ 
        usr.schedule?.wednesday==="true" && dayCount.Wednesday++
        usr.schedule?.thursday==="true" && dayCount.Thursday++
        usr.schedule?.friday==="true" && dayCount.Friday++
      });


      // check which day has largest count 

      //TS prob1
     // cant user Object.values due to TS issue set he lib to es2017. PROBLME TS
  
      let dayName:string="";
      let countNumber =  Math.max(  ...Object.values(dayCount) )
    

     Object.keys(dayCount).filter( (day:string) => dayCount[day] === countNumber).forEach( (day)=>dayName+=` ${day}`)

     this.setState( {        
      ...this.state,
        popularDay:dayName,
        pupularDayCount:countNumber,
        user:{
          displayName:this.context.pageContext.user.displayName,
          email:this.context.pageContext.user.email,
          team: val
        }
      })

      this.renderCard();
    })

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());

    this.quickViewNavigator.register(VIEW_QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    this.quickViewNavigator.register(EDIT_QUICK_VIEW_REGISTRY_ID, () => new EditView());

    return Promise.resolve(); 
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo2.png');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'TeamMeetingDay-property-pane'*/
      './TeamMeetingDayPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.TeamMeetingDayPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {

    //console.log(this.state.user)

    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}

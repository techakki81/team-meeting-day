import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments, AdaptiveCardExtensionContext } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TeamMeetingDayAdaptiveCardExtensionStrings';
import { ITeamMeetingDayAdaptiveCardExtensionProps, ITeamMeetingDayAdaptiveCardExtensionState } from '../TeamMeetingDayAdaptiveCardExtension';

import { weekNumber } from '../../../utilities/weekNumber';

import "@pnp/common"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import {ICamlQuery} from "@pnp/sp/lists"


export interface IQuickViewData {
  subTitle: string;
  title: string; 

  togMon:string;
  togTue:string;
  togWed:string;
  togThu:string;
  togFri:string;
}

export class EditView extends BaseAdaptiveCardView<
  ITeamMeetingDayAdaptiveCardExtensionProps,
  ITeamMeetingDayAdaptiveCardExtensionState,
  IQuickViewData
> {

  
  private async getListItem(){

     const qry:ICamlQuery = {
      ViewXml:`<View><Query><Where><And>
      <Eq><FieldRef Name='Title' /><Value Type='Text'>${this.context.pageContext.user.email}</Value></Eq>
      <Eq><FieldRef Name='Week' /><Value Type='Number'>${weekNumber( new Date() )}</Value></Eq></And></Where></Query></View>`
     }
  
  const lstItms = await this.state.spList.getItemsByCAMLQuery(qry)
  let lstItm = undefined;

  if(lstItms?.length<1)
    {
      const result  = await this.state.spList.items.add( {
              Title: this.context.pageContext.user.email,
              Week: weekNumber(new Date())
            })
            
      lstItm = await result.data
    }
    else  
    {
      lstItm = lstItms[0]     
    }

   this.setState({
     ...this.state,
     spListItem: lstItm
   })

  }
  
  public  get data(): IQuickViewData {

    if(!this.state.spListItem){    
      setTimeout( async() =>{        
        await this.getListItem()      
      },0)
      return undefined
    }
   
    //TOTALK: card title is the first header...you cant change it 
    // get the data for the user...if no data then add one item and return null for all 
    //otherwise populate 
    //TOLKA 
    // the toggle works with on and off values and not tru and fasle. IE you have to define 
    // what is on value and what is off value

    return  {
      title: strings.EditViewTitle, 
      subTitle: "strings.SubTitle ",
      togMon:  String(this.state.spListItem?.Monday),  
      togTue:  String( this.state.spListItem?.Tuesday),
      togWed:  String(this.state.spListItem?.Wednesday),
      togThu:  String(this.state.spListItem?.Thursday),
      togFri:  String(this.state.spListItem?.Friday)
   };   
    
  }   

  public onAction(action: IActionArguments): void {
    
    //TOTALK
    // only one button so no need to check action.type
    let data = (action as ISubmitActionArguments).data  

    console.log(data)

     this.getListItem().then(lstItm =>{
      this.state.spList.items.getById(this.state.spListItem.Id).update({
        Monday:data.togMon || this.state.spListItem.Monday,
        Tuesday:data.togTue|| this.state.spListItem.Tuesday,
        Wednesday:data.togWed|| this.state.spListItem.Wednesday,
        Thursday:data.togThu|| this.state.spListItem.Thursday,
        Friday:data.togFri|| this.state.spListItem.Friday
       })            
     })  

     // close the quick View 
     this.quickViewNavigator.close()
  }   

  public get template(): ISPFxAdaptiveCard {
    return require('./template/EditViewTemplate.json');
  }
}
import { AdaptiveCardExtensionContext } from '@microsoft/sp-adaptive-card-extension-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {ISchedule, IUser} from '../adaptiveCardExtensions/teamMeetingDay/TeamMeetingDayAdaptiveCardExtension'
import * as sampleJson from './sampleUserData.json';
import { MSGraphClient } from '@microsoft/sp-http';
import "@pnp/common"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import {ICamlQuery, IList} from "@pnp/sp/lists"
import { weekNumber } from './weekNumber';

//  export function getUserTeamData(context:AdaptiveCardExtensionContext,spList:IList):Promise<IUser[]> {

//      // if its hosted in local workbench then send the file otherwise make http class
//      // if its workbench then return the normal user
//      //otherwise make a call to graph

//       return new Promise<IUser[]>( (resolve,reject) =>
//       {
     
    
//             if(context.pageContext.site.serverRequestPath.indexOf("/_layouts/workbench.aspx"))
//             import ('./sampleUserData.json').then( (resp:any) => {
                
//                 let users =  resp.value.map( (val:any) =>{

//                     Promise.resolve( getUserSchedule( val.mail, spList)).then( (sch:ISchedule) =>{

//                         console.log(`promise is resolved`)
//                         console.log(sch)
//                         return {
//                         displayName: val.displayName,
//                         email:val.mail,
//                         schedule: sch
//                         }
//                     })
//                 })   
//                 // console.log( users)
//                 resolve(users)
//                 })
//             else 
//             this.context.msGraphClientFactory.getClient().then((client:MSGraphClient):void=>{
//                 client.api("/me/directReports")
//                 .get( (error, response: any,rawResPonse?:any) =>{
                    
//                     let users = response.value.map( (val:any) => <IUser> {
//                         displayName: val.displayName,
//                         email:val.mail,
//                         schedule:getUserSchedule( val.mail, spList)
//                     })

//                 resolve(users)
//                 })
//             })        
//       })
//     }


export  async function getUserTeamData(context:AdaptiveCardExtensionContext,spList:IList): Promise<IUser[]> {

         // if its hosted in local workbench then send the file otherwise make http class
         // if its workbench then return the normal user
         //otherwise make a call to graph
     
        
        if(context.pageContext.site.serverRequestPath.indexOf("/_layouts/workbench.aspx")) {

                    let resp:any = await import ('./sampleUserData.json')
            
                    let users = resp.value.map(  async (val:any) => {

                        let sch:any =  await getUserSchedule( val.mail, spList)

                        return <IUser>{
                            displayName: val.displayName,
                            email:val.mail,
                            schedule: sch
                        }
                    })

                    return  await Promise.all(users)
                }
        else 
        {
            return null
        }

                
                // else 
                // this.context.msGraphClientFactory.getClient().then((client:MSGraphClient):void=>{
                //     client.api("/me/directReports")
                //     .get( (error, response: any,rawResPonse?:any) =>{
                        
                //         let users = response.value.map( (val:any) => <IUser> {
                //             displayName: val.displayName,
                //             email:val.mail,
                //             schedule:getUserSchedule( val.mail, spList)
                //         })
    
                //     resolve(users)
                //     })
                // })        
          
        }

  export async function getUserSchedule( usrEmail:string,  spList:IList, date?:Date):Promise<ISchedule> {   

    let weekNo:number =  weekNumber( date || new Date())

        const qry:ICamlQuery = {
         ViewXml:`<View><Query><Where><And>
         <Eq><FieldRef Name='Title' /><Value Type='Text'>${usrEmail}</Value></Eq>
         <Eq><FieldRef Name='Week' /><Value Type='Number'>${weekNo}</Value></Eq></And></Where></Query></View>`
        }

        const lstItms = await spList.getItemsByCAMLQuery(qry)

        if(lstItms.length>0)
            {
                

                const sch:ISchedule = {
                    monday:lstItms[0].Monday.toString(),
                tuesday:lstItms[0].Tuesday.toString(),
            wednesday:lstItms[0].Wednesday.toString(),
            thursday:lstItms[0].Thursday.toString(),
            friday:lstItms[0].Friday.toString()
        } 
        return sch     
    }
     
     else
     return null

     }     

     
    //  const lstItms =  Promise.resolve(spList.getItemsByCAMLQuery(qry) ).then ( (lstItms:any)=>{

    //     if(lstItms.length>0)
    //  {
    //      let o ={
    //         monday:lstItms[0].Monday,
    //         tuesday:lstItms[0].Tuesday,
    //         wednesday:lstItms[0].Wednesday,
    //         thursday:lstItms[0].Thursday,
    //         friday:lstItms[0].Friday
    //     } 
    //     console.log(o)
    //     return  o
     
    // }
     
    //  else
    //  return null

    //  })     
    
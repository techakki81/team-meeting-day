export function weekNumber(currentdate:Date){
    
var oneJan = new Date(currentdate.getFullYear(),0,1);
let dateDiff:number = (currentdate.valueOf() - oneJan.valueOf()) 
var numberOfDays = Math.floor((dateDiff) / (24 * 60 * 60 * 1000));
var result = Math.ceil(( currentdate.getDay() + 1 + numberOfDays) / 7);
return result

}
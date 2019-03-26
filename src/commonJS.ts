import pnp from 'sp-pnp-js';
import {Web} from 'sp-pnp-js';
import { sp } from "sp-pnp-js";
import 'jquery';
declare var $;
// ADD NEW ITEM

async function addItems(listName: string, listColumns: any) {
  await pnp.sp.web.lists.getByTitle(listName).items.add(listColumns);
}

// ADD NEW ITEM WITH DOCUMENT

async function additemsattachment(listName: string, file: any, listColumns?: any) {
  var result:any;
  result = await pnp.sp.web.getFolderByServerRelativeUrl(listName).files.add(file.name, file, true);
  result.file.listItemAllFields.get().then((listItemAllFields) => {
    pnp.sp.web.lists.getByTitle(listName).items.getById(listItemAllFields.Id).update(listColumns);
  });
 
}

// ADD NEW ITEM WITH IMAGE

async function additemsimage(listName: string, filename:string,file: any, listColumns: any) {
  var result = await pnp.sp.web.getFolderByServerRelativeUrl(listName).files.add(filename, file, true)
  .then(function (result) {
  result.file.listItemAllFields.get().then((listItemAllFields) => {
  return pnp.sp.web.lists.getByTitle(listName).items.getById(listItemAllFields.Id).update(listColumns);
  });
  });
  return result;
}


// READ ITEMS

async function readItems(listName: string, listColumns: string[], topCount: number, orderBy: string, filterKey?: string, filterValue?: any) {
  var matchColumns = formString(listColumns);
  var resultData: any;
  if (filterKey == undefined) {
     resultData = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).top(topCount).orderBy(orderBy,false).get()
  }
  else {
     resultData = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy,false).get()
  }
  return (resultData);
}

// READ Single ITEMS with Lookup

async function readItem(listName: string, listColumns: string[], topCount: number, orderBy: string, filterKey?: string, filterValue?: any, Lookupvalue?: string) {
  var matchColumns = formString(listColumns);
  var resultData: any;
  if(Lookupvalue != "")
  {
      return pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).expand(Lookupvalue).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy,false).get()
  }
  else if (filterKey == undefined) {
     resultData = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).top(topCount).orderBy(orderBy,false).get()
  }
  else {
     resultData = await pnp.sp.web.lists.getByTitle(listName).items.select(matchColumns).filter("" + filterKey + " eq '" + filterValue + "'").top(topCount).orderBy(orderBy,false).get()
  }
  return (resultData);
}
// UPDATE ITEM

async function updateItem(listName: string, id: number, listColumns: any) {
  var result: any;
  result = await pnp.sp.web.lists.getByTitle(listName).items.getById(id).update(listColumns);
  return(result);
}

// DELETE ITEM

async function deleteItem(listName: string, itemID: number) {
  return await pnp.sp.web.lists.getByTitle(listName).items.getById(itemID).delete();
}

// BATCH DELETE - NOT YET TESTED


// async function batchDelete(listName: string, selectedArray: number[]) {  
  
//   let batch = sp.web.createBatch();
//   var arrayLen = selectedArray.length;  
//   for (var i =0; i<arrayLen;i++){
//     //await sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).inBatch(batch).delete().then(r => {
//     await sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).delete().then(r => {
//       console.log("deleted");
//     });
//   }
//   batch.execute().then(() => 
//   location.reload());
//   }

// GET FOLDER ONLY

async function GetFolder(listName: string){
  let folderList = await pnp.sp.web.folders.getByName(listName).folders.expand('ListItemAllFields').get();
  return folderList;
}


// // REMOVE FOLDER FROM DOC LIB

// async function DeleteFolder(listName: string, folderName: string){
//   console.log('common');
//   let confirm= await pnp.sp.web.folders.getByName(listName).folders.getByName(folderName).delete();
//   console.log(confirm);
//   return confirm;
// }
async function batchDelete(listName: string, selectedArray: number[]) {

  let batch = sp.web.createBatch();
  var arrayLen = selectedArray.length;
  for (var i = 0; i < arrayLen; i++) {
    await sp.web.lists.getByTitle(listName).items.getById(selectedArray[i]).delete().then(r => { });
  }
  batch.execute().then(() => location.reload());
}
// REMOVE FOLDER FROM DOC LIB

async function DeleteFolder(listName: string, folderName: string){
  return await pnp.sp.web.folders.getByName(listName).folders.getByName(folderName).delete();
}

// CHECK USER IN GROUP

async function checkUserinGroup(Componentname: string, email: string, callback) {
  var myitems: any[];
  pnp.sp.web.siteUsers
      .getByEmail(email)
      .groups.get()
      .then((items: any[]) => {
          var currentComponent = Componentname;
          myitems = $.grep(items, function (obj, index) {
              if (obj.Title.indexOf(currentComponent) != -1) {
                  return true;
              }
          });
          callback(myitems.length);
      });
}

// GET ALL SUBSITES
 
async function getListOfSubSites( url : string) {
  // var result: any;
  // result = await pnp.sp.web.webs.select().get();
  // return result;
  var result: any;
  let my_web = new Web(url);
  // let batch = web.createBatch();
  result = await my_web.webs.select().get();
  return result;
}
 
// GET LIST OF DOC LIBS
 
async function getListOfDocLib(topCount: number, orderBy: string) {
  var result: any;
  result = await pnp.sp.web.lists.filter('BaseTemplate eq 101').top(topCount).orderBy(orderBy).get();
  return result;
}
 
// FORM STRING

function formString(listColumns: string[]) {
  var formattedString: string = "";
  for (let i = 0; i < listColumns.length; i++) {
    formattedString += listColumns[i] + ',';
  }
  return formattedString.slice(0, -1);
}

// FORMAT DATE

function formatDate(dateVal) {
  var date = new Date(dateVal);
  var year = date.getFullYear();
  var locale = "en-us";
  var month = date.toLocaleString(locale, { month: "long" });
  var dt = date.getDate();
  var dateString: string;
  if (dt < 10) {
      dateString = "0" + dt;
  }
  else
      dateString = dt.toString();
  return dateString + ' ' + month.substr(0, 3) + ',' + year
}

function GetQueryStringParams(sParam) {
  var sPageURL = window.location.search.substring(1);
  var sURLVariables = sPageURL.split('&');
  for (var i = 0; i < sURLVariables.length; i++) {
      var sParameterName = sURLVariables[i].split('=');
      if (sParameterName[0] == sParam) {
          return sParameterName[1];
      }
  }
}

function base64ToArrayBuffer(base64) {
  var binary_string =  window.atob(base64);
  var len = binary_string.length;
  var bytes = new Uint8Array( len );
  for (var i = 0; i < len; i++)        {
      bytes[i] = binary_string.charCodeAt(i);
  }
  return bytes.buffer;
}


export {getListOfDocLib,getListOfSubSites,addItems,readItems,readItem,additemsimage,deleteItem,updateItem,DeleteFolder,GetFolder,formString,additemsattachment,checkUserinGroup,batchDelete,formatDate,GetQueryStringParams,base64ToArrayBuffer}
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/

import { showMessage } from "./message-helper";
import * as $ from "jquery";

export async function callGetUserData(middletierToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/getuserdata`,
      headers: { Authorization: "Bearer " + middletierToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function callCheckGroup(middletierToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/checkGroup`,
      headers: { Authorization: "Bearer " + middletierToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function currentUserDetails(middletierToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/currentUserDetails`,
      headers: { Authorization: "Bearer " + middletierToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function callGetAllADUsers(middletierToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/getAllADUsers`,
      headers: { Authorization: "Bearer " + middletierToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function callGetSearchedADUser(middletierToken: string, searchText: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/getSearchedADUsers`,
      headers: { Authorization: "Bearer " + middletierToken, data: searchText },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function GetAllSiteUsers(middletierToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/GetAllSiteUsers`,
      headers: { Authorization: "Bearer " + middletierToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function GetSiteUserFromEmail(middletierToken: string, Useremail: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/GetSiteUserFromEmail`,
      headers: { Authorization: "Bearer " + middletierToken, Useremail: Useremail },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function GetSPDoc(middletierToken: string, docUrl: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/GetSPDoc`,
      headers: { Authorization: "Bearer " + middletierToken, docurl: docUrl },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function GetSPData(middletierToken: string, searchText: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/GetSPData`,
      headers: { Authorization: "Bearer " + middletierToken, data: searchText },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function CreateRequest(middletierToken: string, listName:string,createItem:any): Promise<any> {
  try {
    const response = await $.ajax({
      type: "POST",
      url: `/CreateRequest`,
      headers: { Authorization: "Bearer " + middletierToken, ListName:listName, Data: JSON.stringify(createItem) },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
export async function UpdateRequest(middletierToken: string, createItem, itemID, listName): Promise<any> {
  try {
    const response = await $.ajax({
      type: "POST",
      url: `/UpdateRequest`,
      headers: { Authorization: "Bearer " + middletierToken, Data: JSON.stringify(createItem), Itemid: `${itemID}`, ListName: listName },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
// export async function DeleteRequest(middletierToken: string, itemID): Promise<any> {
//   try {
//     const response = await $.ajax({
//       type: "POST",
//       url: `/DeleteRequest`,
//       headers: { Authorization: "Bearer " + middletierToken, Itemid: `${itemID}` },
//       cache: false,
//     });
//     return response;
//   } catch (err) {
//     showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
//     throw err;
//   }
// }
export async function GetListData(middletierToken: string,listName: string,selectStr:string,expandStr:string, filterStr: string ): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: `/GetListData`,
      headers: { Authorization: "Bearer " + middletierToken,listName:listName,selectStr:selectStr,expandStr:expandStr, filterStr: filterStr },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}

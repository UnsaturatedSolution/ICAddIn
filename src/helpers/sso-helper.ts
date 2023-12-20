/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthdialog";
import { callGetUserData, DeleteRequest, GetSPData, GetSPDoc, UpdateRequest, GetListData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";
import { callGetSearchedADUser, CreateRequest } from "./middle-tier-calls";
/* global OfficeRuntime */

let retryGetMiddletierToken = 0;
let _middletierToken: string = "";
let _mfaMiddletierToken: string = "";

export async function getUserData(callback): Promise<void> {
  try {
    let middletierToken: string = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response: any = await callGetUserData(middletierToken);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken: string = await OfficeRuntime.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = callGetUserData(mfaMiddletierToken);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

function handleAADErrors(response: any, callback: any): void {
  // On rare occasions the middle tier token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired middle tier token.

  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
    retryGetMiddletierToken++;
    getUserData(callback);
  } else {
    dialogFallback(callback);
  }
}

export async function getSearchUser(querytext: string, callback): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await callGetSearchedADUser(middletierToken, querytext);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      let mfaMiddletierToken: string =
        _mfaMiddletierToken !== ""
          ? _mfaMiddletierToken
          : await OfficeRuntime.auth.getAccessToken({
            authChallenge: response.claims,
          });
      response = callGetSearchedADUser(mfaMiddletierToken, querytext);
    }
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}
export async function GetSPDocSSO(fileURL: string, callback): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await GetSPDoc(middletierToken, fileURL);
    return response;
  } catch (exception) {
  }
}
export async function GetSPDataSSO(querytext: string, callback): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await GetSPData(middletierToken, querytext);
    return response;
  } catch (exception) {
    // if (exception.code) {
    //   if (handleClientSideErrors(exception)) {
    //     dialogFallback(callback);
    //   }
    // } else {
    //   showMessage("EXCEPTION: " + JSON.stringify(exception));
    //   throw exception;
    // }
  }
}
export async function CreateRequestSSO(createItem): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await CreateRequest(middletierToken, createItem);
    return response;
  } catch (exception) {
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}
export async function UpdateRequestSSO(createItem, itemID): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await UpdateRequest(middletierToken, createItem, itemID);
    return response;
  } catch (exception) {
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}
export async function DeleteRequestSSO(itemID): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await DeleteRequest(middletierToken, itemID);
    return response;
  } catch (exception) {
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}
export async function GetSPListData(querytext: string, listname: string, callback): Promise<void> {
  try {
    let middletierToken: string =
      _middletierToken !== ""
        ? _middletierToken
        : await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
    let response: any = await GetListData(middletierToken, querytext, listname);
    return response;
  } catch (exception) {
    console.log(exception)
  }
}

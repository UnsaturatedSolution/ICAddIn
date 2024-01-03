import * as spauth from "node-sp-auth";
import * as request from "request-promise";
import * as appConst from "../constants/appConst";


export async function getSharepointData(req: any, res: any, next: any) {
  spauth
    .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      headers["Accept"] = "application/json;odata=verbose";
      request
        .get({
          url: `${appConst.siteUrl}_api/web`,
          headers: headers,
        })
        .then((response) => {
          console.log(response);
          res.send(response);
        });
    });
}
export async function getAllSiteUsers(req: any, res: any, next: any) {
  spauth
    .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      headers["Accept"] = "application/json;odata=verbose";
      request
        .get({
          url:`${appConst.siteUrl}_api/web/SiteUsers?$select=*,Id&$top=5000`,
          headers: headers,
        })
        .then((response) => {
          console.log(response);
          res.send(response);
        });
    });
}
export async function getSiteUserFromEmail(req: any, res: any, next: any) {
  let userEmail = req.get("Useremail");
  spauth
    .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      headers["Accept"] = "application/json;odata=verbose";
      request
        .get({
          url:`${appConst.siteUrl}_api/web/SiteUsers/getByEmail('${userEmail}')?$select=*,Id`,
          headers: headers,
        })
        .then((response) => {
          console.log(response);
          res.send(response);
        });
    });
}
export async function getSharepointDdoc(req: any, res: any, next: any) {
  let docUrl = req.get("Docurl");
  spauth
    .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      headers["Accept"] = "application/json;odata=verbose";
      request
        .get({
          url: `${appConst.siteUrl}_api/Web/GetFileByServerRelativePath(decodedurl='${docUrl}')`,
          headers: headers,
        })
        .then((response) => {
          console.log(response);
          res.send(response);
        });
    });
}

export async function CreateRequestSP(req, res,next) {
  let data = req.get("Data");
  let listName = req.get("ListName");
  console.log(data);

  spauth
    .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      (headers["Accept"] = "application/json;odata=nometadata"),
        (headers["Content-Type"] = "application/json"),
        request
          .post({
            url: `${appConst.siteUrl}_api/web/lists/getbytitle('${listName}')/items`,
            headers: headers,
            body: data,
          })
          .then((response) => {
            console.log(response);
            res.send(response);
          });
    });
}
export async function UpdateRequestSP(req, res, next) {
  let data = req.get("Data");
  let itemID = req.get("Itemid");
  let listName = req.get("ListName");
  console.log(data);

  spauth
    .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      (headers["Accept"] = "application/json;odata=nometadata"),
        (headers["Content-Type"] = "application/json"),
        (headers["IF-MATCH"] = "*"),
        (headers["X-HTTP-Method"] = "MERGE"),
        request
          .post({
            url: `${appConst.siteUrl}_api/web/lists/getbytitle('${listName}')/items(${itemID})`,
            headers: headers,
            body: data,
          })
          .then((response) => {
            console.log(response);
          });
    });
}
export async function GetListData(req: any, res: any) {
  try {
    /*  console.log('req' + JSON.stringify(req) + 'res' + JSON.stringify(res)); */
    let listName = req.get("listName");
    let selectStr = req.get("selectStr");
    let expandStr = req.get("expandStr");
    let filterStr = req.get("filterStr");
    let urlStr = `${appConst.siteUrl}_api/web/lists/getbyTitle('${listName}')/items?$select=${selectStr}${expandStr != ""?`&$expand=${expandStr}`:""}${filterStr != ""?`&$filter=${filterStr}`:""}`;
    spauth
      .getAuth(`${appConst.siteUrl}`, appConst.spAuthInfo)
      .then((options) => {
        //perform request with any http-enabled library (request-promise in a sample below):
        let headers = options.headers;
        headers["Accept"] = "application/json;odata=verbose";
        request
          .get({
            url:urlStr,
            headers: headers,
            method: "GET",
          })
          .then((response) => {
            // console.log('Get List Data in Sharepoint-helper' + response);

            res.send(response);
          });
      });
  }
  catch (err) {
    console.log(err);
    throw err;
  }
}

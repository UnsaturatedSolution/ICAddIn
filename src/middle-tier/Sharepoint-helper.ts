import * as spauth from "node-sp-auth";
import * as request from "request-promise";

export async function getSharepointData(req: any, res: any, next: any) {
  spauth
    .getAuth("https://vichitra.sharepoint.com/sites/dev/", {
      clientId: "6dc9bad4-bea5-46f1-a47e-6430d2c83ae4",
      clientSecret: "qc0syH8nllqvMlKON7wmR9S8fJ2k/GTVJXqv2PiS0Vc=",
      realm: "6621c5f1-da8a-4ee5-9708-7b6b5334d53b",
    })
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      headers["Accept"] = "application/json;odata=verbose";
      request
        .get({
          url: "https://vichitra.sharepoint.com/sites/dev/_api/web",
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
    .getAuth("https://vichitra.sharepoint.com/sites/dev/", {
      clientId: "6dc9bad4-bea5-46f1-a47e-6430d2c83ae4",
      clientSecret: "qc0syH8nllqvMlKON7wmR9S8fJ2k/GTVJXqv2PiS0Vc=",
      realm: "6621c5f1-da8a-4ee5-9708-7b6b5334d53b",
    })
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      headers["Accept"] = "application/json;odata=verbose";
      request
        .get({
          url: `https://vichitra.sharepoint.com/sites/dev/_api/Web/GetFileByServerRelativePath(decodedurl='${docUrl}')`,
          headers: headers,
        })
        .then((response) => {
          console.log(response);
          res.send(response);
        });
    });
}

export async function CreateRequestSP(req, next, res) {
  let data = req.get("Data");
  console.log(data);

  spauth
    .getAuth("https://vichitra.sharepoint.com/sites/dev/", {
      clientId: "6dc9bad4-bea5-46f1-a47e-6430d2c83ae4",
      clientSecret: "qc0syH8nllqvMlKON7wmR9S8fJ2k/GTVJXqv2PiS0Vc=",
      realm: "6621c5f1-da8a-4ee5-9708-7b6b5334d53b",
    })
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      (headers["Accept"] = "application/json;odata=nometadata"),
        (headers["Content-Type"] = "application/json"),
        request
          .post({
            url: "https://vichitra.sharepoint.com/sites/dev/_api/web/lists/getbytitle('InvestcorpDocumentAssignees')/items",
            headers: headers,
            body: data,
          })
          .then((response) => {
            console.log(response);
          });
    });
}
export async function UpdateRequestSP(req, next, res) {
  let data = req.get("Data");
  let itemID = req.get("Itemid");
  console.log(data);

  spauth
    .getAuth("https://vichitra.sharepoint.com/sites/dev/", {
      clientId: "6dc9bad4-bea5-46f1-a47e-6430d2c83ae4",
      clientSecret: "qc0syH8nllqvMlKON7wmR9S8fJ2k/GTVJXqv2PiS0Vc=",
      realm: "6621c5f1-da8a-4ee5-9708-7b6b5334d53b",
    })
    .then((options) => {
      //perform request with any http-enabled library (request-promise in a sample below):
      let headers = options.headers;
      (headers["Accept"] = "application/json;odata=nometadata"),
        (headers["Content-Type"] = "application/json"),
        (headers["IF-MATCH"] = "*"),
        (headers["X-HTTP-Method"] = "MERGE"),
        request
          .post({
            url: `https://vichitra.sharepoint.com/sites/dev/_api/web/lists/getbytitle('InvestcorpDocumentAssignees')/items(${itemID})`,
            headers: headers,
            body: data,
          })
          .then((response) => {
            console.log(response);
          });
    });
}


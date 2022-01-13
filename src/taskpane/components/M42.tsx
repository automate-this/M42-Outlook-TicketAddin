import { TicketTypes, DefaultCategoryGUID, DefaultInitiatorGUID } from "./Constants";

export async function FetchToken(BaseURL :string, ApiToken :string) {
  const data = fetch(BaseURL + "/api/ApiToken/GenerateAccessTokenFromApiToken/", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + ApiToken,
    },
  })
    .then((response) => response.json())
    .then((responseData) => responseData.RawToken)
    .catch((error) => (document.getElementById("app-message").innerHTML = "Error: " + error));
  return data;
}

export async function FetchFragments(BaseURL :string, AccessToken :string, Fragment :string) {
  const data = fetch(BaseURL + "/api/data/fragments/" + Fragment, {
    method: "get",
    headers: {
      Accept: "application/json",
      Authorization: "Bearer " + AccessToken,
    },
  })
    .then((response) => response.json())
    .catch((error) => (document.getElementById("app-message").innerHTML = "Error: " + error));
  return data;
}

export async function FetchCategories(BaseURL :string, AccessToken :string, searchFilter :string) {
  const data = fetch(BaseURL + `/api/entity/fragments/?entityClass=SPSScCategoryClassBase&parent=Parent&includeRoot=true&searchFilter=${searchFilter}&top=6&notStrictAutocomplete=true`, {
    method: "get",
    headers: {
      Accept: "application/json",
      Authorization: "Bearer " + AccessToken,
    },
  })
    .then((response) => response.json())
    .catch((error) => (document.getElementById("app-message").innerHTML = "Error: " + error));
  return data;
}

export async function FetchTicketType(BaseURL :string, AccessToken :string, ticketNumber :string) {
  const data = fetch(BaseURL + `/api/entity/fragments?entityclass=SPSActivityClassBase&searchFilter=${ticketNumber}`, {
    method: "get",
    headers: {
      Accept: "application/json",
      Authorization: "Bearer " + AccessToken,
    },
  })
    .then((response) => response.json())
    .catch((error) => (document.getElementById("app-message").innerHTML = "Error: " + error));
  return data;
}

export async function CreateTicket(BaseURL: string, AccessToken: string, _Subject: string, _Description: string, _Initiator: string, _TicketType: string, _Transform: boolean, _TransformTo: string, _Recipient: string, _Category: string) {
  let cat = _Category ? _Category : DefaultCategoryGUID;

  //console.log("CreateTicket | _Description: " + _Description);
  
  let InitialData = {};
  if(_TicketType == TicketTypes.Ticket && _Transform){
    //Ticket-Anlage mit direkter Umwandlung: (ganz unten)
    //https://help.matrix42.com/030_DWP/030_INT/Business_Processes_and_API_Integrations/Public_API_reference_documentation/Object_Data_Service%3A_Create_Object
      InitialData = {
        Configuration:{
          TicketType: _TransformTo
        }
      }
  }

  const body = {
    SPSActivityClassBase: {
      Subject: _Subject,
      Category: cat,
      DescriptionHTML: _Description,
      Initiator: _Initiator ? _Initiator : DefaultInitiatorGUID,
      Recipient: _Recipient,
      //EntryBy: 4,
      //NotificationMode: 3, //Never
      NotificationMode: 2, //Creation and Closing
      //NotificationMode: 1, //Always
    },
    InitialData
  };
  const data = fetch(BaseURL + `/api/data/objects/${_TicketType}`, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + AccessToken,
    },
    body: JSON.stringify(body),
  })
    .then((response) => response.json())
    .catch((error) => (document.getElementById("app-message").innerHTML = "Error: " + error));
  return data;
}

export async function UploadFile(BaseURL :string, AccessToken :string, _TicketID :string, _FileName :string, _File: string, _FileType: string, _TicketType: string) {
  const base64 = 'data:' + _FileType + ';base64,' + _File
  const blob = await fetch(base64)
  .then(res => res.blob())

  fetch(
    BaseURL +
      `/api/filestorage/add?entity=${_TicketType}&objectId=${_TicketID}&fileName=${_FileName}`,
    {
      method: "post",
      headers: {
        "Content-Type": "undefined",
        Authorization: "Bearer " + AccessToken,
      },
      body: blob,
    }
  )
  .catch((error) => (document.getElementById("app-message").innerHTML = "Error: " + error));
}

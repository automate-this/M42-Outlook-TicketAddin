import * as React from "react";
import { TextField, ComboBox, Label, Link, MessageBar, MessageBarType, Checkbox, ChoiceGroup, TagPicker, IBasePickerSuggestionsProps } from "office-ui-fabric-react";
import { Spinner, SpinnerSize, DefaultButton, PrimaryButton } from "office-ui-fabric-react";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";


// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-64.png";
import "../../../assets/icon-80.png";
import "../../../assets/icon-128.png";
import "../../../assets/logo-filled.png";
import { FetchToken, FetchFragments, CreateTicket, UploadFile, FetchCategories, FetchTicketType } from "./M42";
import Progress from "./Progress";

import { IAppState } from "./IAppState";
import { IAppProps } from "./IAppProps";
import { TicketTypes, CustomPropertyNames } from "./Constants";
import SuggestionTag from "./SuggestionTag";


const SuggestedItem: (documentProps, itemProps) => JSX.Element = (
  documentProps: any,
  itemProps: any,
) => {

  return (
    <div className="ms-TagItem-TextOverflow" style={{textAlign: "left", maxWidth: 260}}>
      <div style={{padding: '6px 12px 0px', fontSize: 10, color: "grey"}}>{documentProps["info"]}</div>
      <div style={{padding: '0px 12px 7px'}}>{documentProps["name"]}</div>
    </div>
  );
};

export default class App extends React.Component<IAppProps, IAppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      sender: "",
      subject: "",
      body: "",
      attachments: [],
      mailboxItem: undefined,
      mailboxItemProps: undefined,
      initiatorUserId: undefined,
      creatorUserId: undefined,
      baseurl: "",
      apitoken: "",
      accesstoken: "",
      ticketid: "",
      ticketnumber: "",
      creatingTicket: false,
      tickettype: TicketTypes.Ticket,
      ticketCategory: "",
      ticketTransform: true,
      ticketTransformTo: "6",
      ticketRealType: "",
      assignToMe: true
    };
    this.handleChange = this.handleChange.bind(this);
    this.saveSettings = this.saveSettings.bind(this);
    this.filterSuggestedTags = this.filterSuggestedTags.bind(this);
  }

  handleChange = (name) => (event) => {
    this.setState({ [name]: event.target.value });
  };
  handleChangeCC = (name, key) => {
    this.setState({ [name]: key });
    console.log("setState " + name + " -> " + key);
  };
  customPropsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      // Handle the failure.
      console.log("customPropsCallback | error");
    } else {
      // Successfully loaded custom properties

      this.setState({mailboxItemProps: asyncResult.value}, () => {
        //Object.keys(mailboxItemProps.rawData).forEach(key => console.log(`customPropsCallback | >> ${key} = ${mailboxItemProps.rawData[key]}`));

        const existingTicketID = this.getMailboxItemProperty(CustomPropertyNames.M42TicketID);
        if(existingTicketID){
          this.setState({ ticketid: existingTicketID, 
                          ticketnumber: this.getMailboxItemProperty(CustomPropertyNames.M42TicketNumber),
                          ticketRealType: this.getMailboxItemProperty(CustomPropertyNames.M42TicketType)},
                          () => {
                            FetchTicketType(this.state.baseurl, this.state.accesstoken, this.state.ticketnumber)
                            .then((data) => {
                              if(data && data.Fragments && data.Fragments[0] && data.Fragments[0]["ObjectId"]==this.state.ticketid){
                                console.log("FetchTicketType | " + data.Fragments[0]["EntityType"]);
                                this.setState({ticketRealType: data.Fragments[0]["EntityType"]})
                              }
                            })
                          })
        }
      });

      
    }
  }
  getMailboxItemProperty(name): any {
    const { mailboxItemProps } = this.state;
    console.log(`getMailboxItemProperty | mailboxItemProps[${name}]: ${mailboxItemProps.rawData[name]}`);
    return mailboxItemProps.rawData[name];
  }
  updateMailboxItemProperty(name, value) {
    const { mailboxItemProps } = this.state;
    mailboxItemProps.set(name, value);
    // Save all custom properties to server.
    mailboxItemProps.saveAsync(this.saveMailboxItemPropsCallback);
  }
  removeMailboxItemProperty(name) {
    const { mailboxItemProps } = this.state;
    mailboxItemProps.remove(name);
    // Save all custom properties to server.
    mailboxItemProps.saveAsync(this.saveMailboxItemPropsCallback);
  }
  saveMailboxItemPropsCallback(result) {
    if (result.status == Office.AsyncResultStatus.Failed) {
      console.log("Action (save mailboxItem property) failed with error: " + result.error.message);
    }else{
      console.log(`MailboxItem Props saved with status: ${result.status}`);
    }
  }
  callback(result) {
    if (result.value.length > 0) {
      for (let i = 0; i < result.value.length; i++) {
        result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, this.handleAttachmentsCallback);
      }
    }
  }

  handleAttachmentsCallback(result) {
    // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
    switch (result.value.format) {
      case Office.MailboxEnums.AttachmentContentFormat.Base64:
        // Handle file attachment.
        let att = {
          id: "",
          name: "",
          type: "",
          file: "",
        };
        att.id = result.asyncContext.ID;
        att.name = result.asyncContext.name;
        att.type = result.asyncContext.contentType;
        att.file = result.value.content;
        this.state.attachments.push(att);
        break;
      case Office.MailboxEnums.AttachmentContentFormat.Eml:
        // Handle email item attachment.
        break;
      case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
        // Handle .icalender attachment.
        break;
      case Office.MailboxEnums.AttachmentContentFormat.Url:
        // Handle cloud attachment.
        break;
      default:
      // Handle attachment formats that are not supported.
    }
  }

  async componentDidMount(): Promise<void> {
    try {
      let bUrl = Office.context.roamingSettings.get("baseurl");
      let apiToken = Office.context.roamingSettings.get("apitoken");
      let AccessToken : string = await FetchToken(bUrl, apiToken);
      this.setState({
        baseurl: bUrl,
        apitoken: apiToken,
        accesstoken: AccessToken
      }, async () => {

        let outlookUserMail = await FetchFragments(this.state.baseurl, this.state.accesstoken,
          "SPSUserClassBase?where=MailAddress='" + Office.context.mailbox.userProfile.emailAddress + "'&columns=ID"
        );

        let mailboxItemLocal = Office.context.mailbox.item;
        mailboxItemLocal.loadCustomPropertiesAsync(this.customPropsCallback.bind(this));

        let initiatorUsers = await FetchFragments(this.state.baseurl, this.state.accesstoken,
          "SPSUserClassBase?where=MailAddress='" + mailboxItemLocal.from.emailAddress + "'&columns=ID"
        );
        
        mailboxItemLocal.body.getAsync(Office.CoercionType.Html, { asyncContext: "This is passed to the callback" },
          function (result: Office.AsyncResult<string>) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.log("Action (get item body) failed with error: " + result.error.message);
            } else {
              this.setState( {body: result.value} );
            }
          }.bind(this)
        );
        for (let i = 0; i < mailboxItemLocal.attachments.length; i++) {
          const ID = mailboxItemLocal.attachments[i].id;
          const name = mailboxItemLocal.attachments[i].name;
          const contentType = mailboxItemLocal.attachments[i].contentType;
          const options = {
            asyncContext: {
              currentItem: mailboxItemLocal,
              ID: ID,
              name: name,
              contentType: contentType,
            },
          };
          mailboxItemLocal.getAttachmentContentAsync(ID, options, this.handleAttachmentsCallback.bind(this));
        }
        
        this.setState({
          mailboxItem: mailboxItemLocal,
          sender: mailboxItemLocal.from.emailAddress,
          subject: mailboxItemLocal.subject,
          initiatorUserId: initiatorUsers.length > 0 ? initiatorUsers[0].ID : undefined,
          creatorUserId: outlookUserMail.length > 0 ? outlookUserMail[0].ID : undefined
        });
      });
      
    } catch (error) {
      console.log(error);
    }
  }

  ticketCreationDisabled(): boolean {
    return Boolean(/*!this.state.initiatorUserId || */this.state.ticketid || this.state.creatingTicket || !this.state.subject || this.state.baseurl==null || this.state.baseurl=="" || this.state.apitoken==null || this.state.apitoken=="");
  }

  saveSettings = () => {
    if (!this.state.baseurl || !this.state.apitoken) {
      document.getElementById("settings-message").style.color = "red";
      document.getElementById("settings-message").innerHTML = "Bitte füllen Sie beide Felder aus.";
    } else {
      Office.context.roamingSettings.set("baseurl", this.state.baseurl);
      Office.context.roamingSettings.set("apitoken", this.state.apitoken);

      Office.context.roamingSettings.saveAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Action (save settings) failed with message ${result.error.message}`);
          document.getElementById("settings-message").style.color = "red";
          document.getElementById("settings-message").innerHTML = "Fehler beim Speichern der Einstellungen";
        } else {
          console.log(`Settings saved with status: ${result.status}`);
          document.getElementById("settings-message").style.color = "green";
          document.getElementById("settings-message").innerHTML = "Einstellungen gespeichert";
        }
      });
    }
  };

  createTicketClick = async () => {
    this.setState({creatingTicket: true});
    const { baseurl, accesstoken, subject, body, attachments, initiatorUserId: userid, tickettype, ticketTransform, ticketTransformTo, assignToMe, creatorUserId, ticketCategory } = this.state;
    
    let modifiedBody: string = body;
    //Modify Body (replace cid-references with the file's base64)
    if (attachments.length > 0) {
      for (let i = 0; i < attachments.length; i++) {
        console.log(`create ticket | replace cid-references | Name: ${attachments[i].name}, Type: ${attachments[i].type}`);
        if(attachments[i].type.indexOf("image") > -1){

          let startingIndexOfSrcAtt:number = body.indexOf(`src="cid:${attachments[i].name}`);//src="cid:image001.png@01D7DA24.65246200"
          console.log(`create ticket | replace cid-references | startingIndexOfSrcAtt of 'src="cid:${attachments[i].name}': ${startingIndexOfSrcAtt}`);
          if(startingIndexOfSrcAtt > -1){
            const base64 = 'data:' + attachments[i].type + ';base64,' + attachments[i].file;
            let endingIndexOfSrcAtt:number = body.indexOf('"', startingIndexOfSrcAtt + 8);
            let origSrcAtt:string = body.substring(startingIndexOfSrcAtt, endingIndexOfSrcAtt + 1);
            console.log(`create ticket | replace cid-references | origSrcAtt: ${origSrcAtt}`);

            modifiedBody = modifiedBody.replace(origSrcAtt, `src="${base64}"`);
          }
        }
      }
    }

    //Create Ticket
    const TicketID = await CreateTicket(baseurl, accesstoken, subject, modifiedBody, userid, tickettype, ticketTransform, ticketTransformTo, assignToMe ? creatorUserId : undefined, ticketCategory);
    const RealTicketType = this.getRealTicketType(tickettype, ticketTransform, ticketTransformTo);
    console.log("create ticket | TicketID: " + TicketID);

    //Upload Attachments
    if (attachments.length > 0) {
      for (let i = 0; i < attachments.length; i++) {
        await UploadFile(baseurl, accesstoken, TicketID, attachments[i].name, attachments[i].file, attachments[i].type, RealTicketType);
        console.log(`create ticket | attachment uploaded (${i+1}/${attachments.length}) `);
      }
    }

    //Get Ticket Number
    let TicketNumbers = await FetchFragments(baseurl,accesstoken,`/${RealTicketType}?where=ID='${TicketID}'&columns=RelatedSPSActivityClassBase.TicketNumber AS TicketNumber`);
    console.log("create ticket | TicketNumber: " + TicketNumbers[0].TicketNumber);
    this.setState({ticketid: TicketID, ticketRealType: RealTicketType, creatingTicket: false, ticketnumber: TicketNumbers[0].TicketNumber});
    try{
      this.updateMailboxItemProperty(CustomPropertyNames.M42TicketID, TicketID);
      this.updateMailboxItemProperty(CustomPropertyNames.M42TicketNumber, TicketNumbers[0].TicketNumber);
      this.updateMailboxItemProperty(CustomPropertyNames.M42TicketType, RealTicketType);
    }catch(e){
      console.log("create ticket | error in updating mailbox item properties | " + e);
    }
  };

  showTicket = async () => {

    console.log("showTicket | platform: " + Office.context.platform);

    let url : string = this.state.baseurl.substring(0, this.state.baseurl.length - 12).substring(this.state.baseurl.indexOf("https")) +
                        `/wm/app-ServiceDesk/notSet/preview-object/${this.state.ticketRealType}/${this.state.ticketid}/0/`;

    if (Office.context.platform.toString().indexOf("OfficeOnline") > -1) {
      window.open(url, "_blank");
    } else {
      Office.context.ui.openBrowserWindow(url);
    }
  }

  getRealTicketType(_selectedType: string, _transforming: boolean, _transformingTo: string): string {
    if(_selectedType == TicketTypes.Ticket && _transforming){
      if(_transformingTo == "0"){
        return TicketTypes.Incident;
      }else if(_transformingTo == "6"){
        return TicketTypes.ServiceRequest;
      }else{
        return _selectedType;
      }
    }else{
      return _selectedType;
    }
  }


  pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: undefined,
    noResultsFoundText: 'Keine Kategorien gefunden',
  };

  async fetchMappedCategories(filterText: string) : Promise<SuggestionTag[]> {
    const { baseurl, accesstoken } = this.state;
    const categories = await FetchCategories(baseurl, accesstoken, filterText);

    const fragments = categories["Fragments"];
    return fragments.map(item => ({ key: item.Id, name: item.DisplayString, info: item.ParentDisplayString })) as SuggestionTag[];
  }

  filterSuggestedTags = async (filterText: string, tagList: SuggestionTag[]): Promise<SuggestionTag[]> => {
    console.log("filterSuggestedTags | " + filterText);

    return filterText
      ? await this.fetchMappedCategories(filterText)
      : [];
  };
  
  getTextFromItem = (item: SuggestionTag) => item.name;


  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome__main">
        <Pivot
          aria-label="M42Pivot"
          styles={{ root: { display: "flex", justifyContent: "center", textAlign: "center" } }}
        >
          <PivotItem
            headerText="Home"
            headerButtonProps={{
              "data-order": 1,
              "data-title": "Home",
            }}
            itemIcon="Home"
            style={{ alignContent: "center" }}
          >
            <Label>Initiator</Label>
            <TextField
              value={this.state.sender}
              borderless
              readOnly
            />
            <TextField
              value={this.state.initiatorUserId}
              placeholder={this.state.sender != "" ? "- (extern)" : undefined}
              borderless
              readOnly
            />
            <TextField
              label="Subject"
              value={this.state.subject}
              onChange={this.handleChange("subject")}
              disabled={this.ticketCreationDisabled() && (this.state.subject != "")}
            />
            <ComboBox
              label="Ticket-Typ"
              selectedKey={this.state.tickettype}
              onChange={(e, option) => {this.handleChangeCC("tickettype", option.key);}}
              options={[{key:TicketTypes.ServiceRequest, text: "Serviceanfrage"},
                        {key:TicketTypes.Incident, text: "Störung"},
                        {key:TicketTypes.Ticket, text: "Ticket"}]}
              disabled={this.ticketCreationDisabled()}
            />
            {this.state.tickettype===TicketTypes.Ticket &&
              <div>
                <Checkbox
                  label="umwandeln in..." 
                  styles={{root: {paddingTop: 10, paddingLeft: 15}}}
                  checked={this.state.ticketTransform}
                  onChange={(e, isChecked) => {this.handleChangeCC("ticketTransform", isChecked);}}
                  disabled={this.ticketCreationDisabled()}/>
                <ChoiceGroup
                  styles={{root: {paddingLeft: 40}}}
                  selectedKey={this.state.ticketTransformTo}
                  onChange={(e, option) => {this.handleChangeCC("ticketTransformTo", option.key);}}
                  options={[{key: "6", text: "Serviceanfrage"},
                            {key: "0", text: "Störung"}]}
                  disabled={!this.state.ticketTransform || this.ticketCreationDisabled()}/>
              </div>
            }

            <Label disabled={this.ticketCreationDisabled()}>Kategorie</Label>
            <TagPicker
              removeButtonAriaLabel="Remove"
              onRenderSuggestionsItem={SuggestedItem as any}
              onResolveSuggestions={this.filterSuggestedTags}
              getTextFromItem={this.getTextFromItem}
              pickerSuggestionsProps={this.pickerSuggestionsProps}
              itemLimit={1}
              onChange={(items?: SuggestionTag[]) => {this.handleChangeCC("ticketCategory", items[0]?.key);}}
              disabled={this.ticketCreationDisabled()}
            />
            
            <Checkbox
                label="mir zuweisen" 
                styles={{root: {paddingTop: 10, paddingLeft: 15}}}
                checked={this.state.assignToMe}
                onChange={(e, isChecked) => {this.handleChangeCC("assignToMe", isChecked);}}
                disabled={this.ticketCreationDisabled()}
            />

            <div className="ms-welcome__main">
              <PrimaryButton
                text="Create Ticket"
                className="ms-welcome__action"
                iconProps={{ iconName: "ChevronRight" }}
                onClick={this.createTicketClick}
                disabled={this.ticketCreationDisabled()}
              />
            </div>
            <div className="ms-welcome__main">
              {this.state.creatingTicket &&
                <Spinner size={SpinnerSize.large} label="Creating ticket"/>
              }
              {this.state.ticketid &&
                <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
                  Ticket created: {this.state.ticketnumber}
                  <br/>
                  {this.state.ticketRealType &&
                    <Link underline
                      onClick={this.showTicket}>
                      Show Ticket
                    </Link>
                  }
                </MessageBar>
              }

                <div style={{display: "none"}}>
                  <br/><br/>
                  <DefaultButton
                    text="Reset Ticket Info (state only)"
                    onClick={() => {this.setState({ticketid: ""});}}
                  />
                </div>
              

              {(this.state.baseurl==null || this.state.baseurl=="" || this.state.apitoken==null || this.state.apitoken=="") &&
                <><br/>
                  <Label style={{color: "FireBrick"}}>Bitte Angaben in den Settings prüfen</Label>
                </>
              }
            </div>
            <p id="app-message" className="ms-font-l" style={{color: "FireBrick"}}></p>
          </PivotItem>
          <PivotItem headerText="Settings" itemIcon="Settings">
            <TextField
              label="M42 ServiceStore Service URL"
              required
              value={this.state.baseurl}
              onChange={this.handleChange("baseurl")}
              placeholder="e.g. https://my.example.com/M42Services"
            />

            <TextField
              label="API Token"
              required
              value={this.state.apitoken}
              onChange={this.handleChange("apitoken")}
              type="password"
            />

            <p id="settings-message" className="ms-font-l"></p>
            <div className="ms-welcome__main">
              <PrimaryButton
                className="ms-welcome__action"
                iconProps={{ iconName: "Save" }}
                onClick={this.saveSettings}
              >
                Save
              </PrimaryButton>
            </div>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}

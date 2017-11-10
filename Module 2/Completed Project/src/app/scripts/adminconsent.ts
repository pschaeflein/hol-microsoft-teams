import { Guid } from "./guid";

/**
* Implementation of the AdminConsent page
*/
export class AdminConsent {
    /**
    * Constructor for Tab that initializes the Microsoft Teams script and themes management
    */
    constructor() {
      microsoftTeams.initialize();
    }

    public requestConsent(tenantId:string) {
      let host = "https://" + window.location.host;
      let redirectUri = "https://" + window.location.host + "/adminconsent.html";
      let clientId = "11dcd3a3-c794-4aec-a7b8-9d66499ed559";
      let state = Guid.NewGuid();
      localStorage.setItem("adminConsent.state",state);

      var consentEndpoint = "https://login.microsoftonline.com/common/adminconsent?" +
                            "client_id=" + clientId +
                            "&state=" + state +
                            "&redirect_uri=" + redirectUri;

      window.location.replace(consentEndpoint);
    }

    public processResponse(response:boolean, error:string){
      if (response) {
        microsoftTeams.authentication.notifySuccess();
      } else {
        microsoftTeams.authentication.notifyFailure(error);
      }
    }
  }
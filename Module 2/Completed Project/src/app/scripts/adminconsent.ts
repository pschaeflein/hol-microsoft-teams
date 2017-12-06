import { Guid } from "./guid";
import { AADAppConfig } from "./aadAppConfig";

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
      let state = Guid.NewGuid();

      var consentEndpoint = "https://login.microsoftonline.com/common/adminconsent?" +
                        "client_id=" + AADAppConfig.clientID +
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
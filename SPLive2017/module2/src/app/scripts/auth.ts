import { UserAgentApplication } from "msal";
import { AADAppConfig } from "./aadAppConfig";

/**
* Implementation of the Auth page
*/
export class Auth {
  private token: string = "";
  private app: UserAgentApplication;
  private user: any;

  /**
  * Constructor for Tab that initializes the Microsoft Teams script
  */
  constructor() {
    microsoftTeams.initialize();

    this.app = new UserAgentApplication(
      AADAppConfig.clientID,
      "https://login.microsoftonline.com/common",
      this.tokenReceivedCallback
    );
  }

  public performAuthV2(level: string) {
    if (this.app.isCallback(window.location.hash)) {
      this.app.handleAuthenticationResponse(window.location.hash);
    }
    else {
      this.user = this.app.getUser();
      if (!this.user) {
        // If user is not signed in, then prompt user to sign in via loginRedirect.
        // This will redirect user to the Azure Active Directory v2 Endpoint
        this.app.loginRedirect(AADAppConfig.graphScopes);
      } else {
        this.getToken();
      }
    }
  }

  private getToken() {
    // In order to call the Graph API, an access token needs to be acquired.
    // Try to acquire the token used to query Graph API silently first:
    this.app.acquireTokenSilent(AADAppConfig.graphScopes).then(
      (token) => {
        //After the access token is acquired, return to MS Teams, sending the acquired token
        microsoftTeams.authentication.notifySuccess(token);
      },
      (error) => {
        // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
        // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user
        // can reenter the current username/ password and/ or give consent to new permissions your application is requesting.
        if (error) {
          this.app.acquireTokenRedirect(AADAppConfig.graphScopes);
        }
      }
    );
  }

  private tokenReceivedCallback(errorDesc, token, error, tokenType) {
    if (token) {
      this.user = this.app.getUser()!;
      microsoftTeams.authentication.notifySuccess(token);
    }
    else {
      microsoftTeams.authentication.notifyFailure(error);
    }
  }
}
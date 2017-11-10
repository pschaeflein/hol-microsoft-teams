# Developing Extensions for Microsoft Teams - Module 2
----------------
In this module, you will extend a generated tab to call the Microsoft Graph. The Exercise contains many code files. The **Completed Project** folder contains files that contain the code and are provided to facilitate copy/paste of the code rather than typing.

## Call Microsoft Graph inside a Tab

This section of the lab contains many code files. The **Lab Files** folder contains files that contain the code and are provided to facilitate copy/paste of the code rather than typing.

1. Open a **Command Prompt** window.
1. Change to the directory where you will create the tab.

  > **Note:** Directory paths can become quite long after node modules are imported.  **We suggest you use a directory name without spaces in it and create it in the root folder of your drive.**  This will make working with the solution easier in the future and protect you from potential issues associated with long file paths. In this example, we use `c:\Dev` as the working directory.

1. Type `md teams-mod2` and press **Enter**.
1. Type `cd teams-mod2` and press **Enter**.

### Run the Yeoman Teams Generator

1. Type `yo teams` and press **Enter**.

1. When prompted:
    1. Accept the default **teams-mod-2** as your solution name and press **Enter**.
    1. Select **Use the current folder** for where to place the files and press **Enter**.
1. The next set of prompts asks for specific information about your Teams App:
    1. Accept the default **teams mod2** as the name of your Microsoft Teams App project and press **Enter**.
    1. Enter your name and press **Enter**.
    1. Enter **https://tbd.ngrok.io** as the URL where you will host this tab and press **Enter**. (We will change this URL later.)
    1. Accept the default selection of **Tab** for what you want to add to your project and press **Enter**.
    1. Enter **Module 2** as the default tab name and press **Enter**.

  >**Note:** At this point, Yeoman will install the required dependencies and scaffold the solution files along with the basic tab. This might take a few minutes.

### Run the ngrok secure tunnel application

1. Open a new **Command Prompt** window.
1. Change to the directory that contains the ngrok.exe application.
1. run the command `ngrok http 3007`
1. The ngrok application will fill the entire prompt window. Make note of the Forwarding address using https. This address is required in the next step.
1. Minimize the ngrok Command Prompt window. It is no longer referenced in this exercise, but it must remain running.

### Register an application in AAD

To enable an application to call the Microsoft Graph, an application registration is required. This lab uses the [Azure Active Directory v2.0 endpoint](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-compare).

1. Open a browser to the url **https://apps.dev.microsoft.com**
1. Log in with a Work or School account.
1. Click **Add an app**
1. Complete the **Register your application** section, entering an Application name. Clear the checkbox for Guided Setup. Click **Create**
1. On the registration page, in the **Platforms** section, click **Add Platform**.
1. In the **Add Platform** dialog, click **Web**.
1. Using the hostname from ngrok, enter a **Redirect URL** to the auth.html file.

    ```
    https://[replace-this].ngrok.io/auth.html
    ```

1. Click the **Add URL** button.
1. Using the hostname from ngrok, enter a **Redirect URL** to the adminconsent.html file.

    ```
    https://[replace-this].ngrok.io/adminconsent.html
    ```
1. Click **Save**.
1. Make note of the Application Id. This value is used in the authentication / token code.
1. Request permission to read Groups.
    1. Scroll down to the **Microsoft Graph Permissions** section.
    1. Next to **Delegated Permissions**, click the **Add** button.
    1. In the **Select Permission** dialog, scroll down and select **Group.Read.All**. Click OK.
    1. Click **Save**.

### Add the Microsoft Authentication Library (MSAL) to the project

Using **npm**, add the Microsoft Authentication library to the project.

1. Open a **Command Prompt** window.
1. Change to the directory containing the tab application.
1. Run the following command:

    ```shell
    npm install msal
    ```

### Configure AAD in project

1. Add a new file to the **scripts** folder named **aadAppConfig.ts**.
1. Add the following to the **aadAppConfig.ts** file. Replace the token enter-your-app-id with the Application Id from the application registration page.

    ```typescript
    export class AADAppConfig {
      static clientID: string = "enter-your-app-id";
      static graphScopes: string[] =  ["https://graph.microsoft.com/user.read", "https://graph.microsoft.com/group.read.all"];
    }
    ```

### Update TypeScript configuration

The MSAL package contains TypeScript Declaration files within the package. This means that the declaration files are contained in the `node-modules` folder. Visual Studio code will discover these declaration files and provide intellisense.

However, the generated TypeScript configuration for client (browser) files will look only in the `src` folder for declaration files. This configuration will cause the build/serve process to fail. Update the TypeScript configuration using the following steps.

1. Open the file **tsconfig-client.js**
1. Locate the `moduleResolution` key in the `compilerOptions` section.
1. Set the value to `"node"`

### Create Tab Configuration page

The Tab in this module can be configured to read information from Microsoft Graph about the current member or about the Group in which the channel exists. Perform the following to create the Tab configuration.

These steps assume that the application created earlier is named **teams-mod-2** and the tab is named **Module2**. Furthermore, paths listed in this section are relative to the `src/app/` folder in the generated application.

1. Open the file **web/module2Config.html**
    1. Locate the `<div>` element with the class of `settings-container`. Replace that element with the following code snippet.

        ```html
        <div class="settings-container">
          <div class="section-caption">Settings</div>
          <div class="form-field-title">
            <div for="graph">Microsoft Graph Functionality:</div>
          </div>
          <div>
            <select name="graph" id="graph"  class="form-control" onchange="onChange(this.value);">
              <option value="" selected>Select one...</option>>
              <option value="member">Member information</option>
              <option value="group">Group information (requires admin consent)</option>
            </select>
          </div>
          <div class="form-field-title">
            <a href="#" onclick="requestConsent();">Provide administrator consent - click if Tenant Admin</a>
          </div>
        </div>
        ```

    1. Add the following function to the `<script>` tag on the page.

        ```js
        function requestConsent() {
          c.getAdminConsent();
          return false;
        }
        ```

1. Open the file **scripts/module2Config.ts**.
    1. Locate the constructor method.  Replace the constructor with the following code snippet. (The snippet includes code to save the **tenantId** for use in the admin consent process.)

        ```typescript
        tenantId?: string;

        constructor() {
          microsoftTeams.initialize();

          microsoftTeams.getContext((context: microsoftTeams.Context) => {
            TeamsTheme.fix(context);
            this.tenantId = context.tid;
            let val = <HTMLInputElement>document.getElementById("graph");
            if (context.entityId) {
              val.value = context.entityId;
            }
            this.setValidityState(val.value !== "");
          });

          microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
            let val = <HTMLInputElement>document.getElementById("graph");

            // Calculate host dynamically to enable local debugging
            let host = "https://" + window.location.host;
            microsoftTeams.settings.setSettings({
              contentUrl: host + "/teamsApp1TabTab.html",
              suggestedDisplayName: 'teamsApp1 Tab',
              removeUrl: host + "/teamsApp1TabRemove.html",
              entityId: val.value
            });

            saveEvent.notifySuccess();

          });
        }
        ```

    1. The tab configuration page has a link for granting admin consent. This link has an `onclick` event. Add the following function to the **module2Configure** object.

        ```typescript
        public getAdminConsent() {
          microsoftTeams.authentication.authenticate({
            url: "/adminconsent.html?tenantId=" + this.tenantId,
            width: 800,
            height: 800,
            successCallback: () => { },
            failureCallback: (err) => { }
          });
        }
        ```

1. Add a new file to the **web** folder named **adminconsent.html**
    1. Add the following to the **adminconsent.html** file.

        ```html
        <!DOCTYPE html>
        <html lang="en">

        <head>
          <meta charset="UTF-8">
          <title>AdminConsent</title>
          <!-- inject:css -->
          <!-- endinject -->
        </head>

        <body>
          <script src="https://statics.teams.microsoft.com/sdk/v1.0/js/MicrosoftTeams.min.js"></script>
          <!-- inject:js -->
          <!-- endinject -->

          <script type="text/javascript">
            function getURLParam(name) {
              var url = window.location.search.substring(1);
              var variables = url.split('&');
              for (var i = 0; i < variables.length; i++) {
                var variable = variables[i].split('=');
                if (variable[0] === name) {
                  return decodeURIComponent(variable[1]);
                }
              }
            }

            var ac = new teamsMod2.AdminConsent();

            var response = getURLParam("admin_consent");
            if (response) {
              ac.processResponse(true);
            } else {
              var error = getURLParam("error_description")
              if (error) {
                ac.processResponse(false, error);
              } else {
                var tenantId = getURLParam("tenantId");
                ac.requestConsent(tenantId);
              }
            }
          </script>
        </body>
        </html>
        ```

1. Add a new file to the **scripts** folder named **adminconsent.ts**
    1. Add the following to the **adminconsent.ts** file.

        ```typescript
        import { Guid } from "./guid";
        import { AADAppConfig } from "./aadAppConfig";

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
        ```

1. Locate the file **scripts/client.ts**
    1. Add the following line to the bottom of **scripts/client.ts**

        ```typescript
        export * from './adminconsent';
        export * from './aadAppConfig';
        export * from './guid';
        ```

1. Following the steps from Exercise 1, redeploy the app. In summary, perform the following steps.
    1. Update the manifest (if necessary) with the ngrok forwarding address.
    1. Run `gulp build` and resolve errors.
    1. Run `gulp serve`
    1. Sideload the app into Microsoft Teams.

1. Add the tab to a channel, or update the settings of the tab in the existing channel. (To update the settings of an existing tab, click the chevron next to the tab name.)

1. Click the **Provide administrator consent - click if Tenant Admin** link.

1. Verify that the Azure Active Directory login and consent flow completes. (If you log in with an account that is not a Tenant administrator, the consent action will fail. Admin Consent is only necessary to view the Group information, not the member information.)

### Content Page and Authentication

With the tab configured, the content page can now render information as selected.  Perform the following to update the Tab content.

1. Open the file **web/module2Tab.html**
    1. Locate the `<div>` element with the id of `app`. Replace that element with the following code snippet.

        ```html
        <div id='app'>
          Loading...
        </div>
        <div>
          <button id="getDataButton">Get MSGraph Data</button>
          <div id="graph"></div>
        </div>
        ```

1. Open the file **scripts/module2Tab.ts**.
    1. Locate the constructor method.  Replace the constructor with the following code snippet. (The snippet includes code to save the configured value as a class-level variable.)

        ```typescript
        configuration?: string;
        groupId?: string;
        token?: string;

        /**
        * Constructor for Tab that initializes the Microsoft Teams script
        */
        constructor() {
          microsoftTeams.initialize();
        }
        ```

    1. Locate the `doStuff` method. Replace the method with the following code snippet. This method will display the configured value and attach a handler to the GetData button.

        ```typescript
        public doStuff() {
          let button = document.getElementById('getDataButton');
          button!.addEventListener('click', e => { this.refresh(); });

          microsoftTeams.getContext((context: microsoftTeams.Context) => {
            TeamsTheme.fix(context);
            this.groupId = context.groupId;
            // hack
            if (context.entityId) {
              this.configuration = context.entityId;
              let element = document.getElementById('app');
              if (element) {
                element.innerHTML = `The value is: ${this.configuration}`;
              }
            }
          });
        }
        ```

    1. Add the following function to the `module2Tab` object. This function runs in response to the button click.

        ```typescript
        public refresh() {
          let graphElement = document.getElementById("graph");
          graphElement!.innerText = "Loading...";
          if (this.token === null) {
            microsoftTeams.authentication.authenticate({
              url: "/auth.html",
              width: 400,
              height: 400,
              successCallback: (data) => {
                // Note: token is only good for one hour
                this.token = data!;
                this.getData(this.token);
              },
              failureCallback: function (err) {
                document.getElementById("graph")!.innerHTML = "Failed to authenticate and get token.<br/>" + err;
              }
            });
          }
          else {
            this.getData(this.token);
          }
        }
        ```

    1. Add the follow method to the `module2Tab` class. This method uses XMLHttp to make a call to the Microsoft Graph and displays the result.

        ```typescript
        public getData(token: string) {
          let graphEndpoint = "https://graph.microsoft.com/v1.0/me";
          if (this.configuration === "group") {
            graphEndpoint = "https://graph.microsoft.com/v1.0/groups/" + this.groupId;
          }

          var req = new XMLHttpRequest();
          req.open("GET", graphEndpoint, false);
          req.setRequestHeader("Authorization", "Bearer " + token);
          req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");
          req.send();
          var result = JSON.parse(req.responseText);
          document.getElementById("graph")!.innerHTML = `<pre>${JSON.stringify(result, null, 2)}</pre>`;
        }
        ```

1. Add a new file to the **web** folder named **auth.html**
    1. Add the following to the **adminconsent.html** file.

        ```html
        <!DOCTYPE html>
        <html lang="en">
        <head>
          <meta charset="UTF-8">
          <title>Auth</title>
          <!-- inject:css -->
          <!-- endinject -->
        </head>
        <body>
          <script src="https://secure.aadcdn.microsoftonline-p.com/lib/0.1.1/js/msal.min.js"></script>
          <script src="https://statics.teams.microsoft.com/sdk/v1.0/js/MicrosoftTeams.min.js"></script>
          <!-- inject:js -->
          <!-- endinject -->
          <script type='text/javascript'>
            var auth = new teamsMod2.Auth();
            auth.performAuthV2();
          </script>
        </body>
        </html>
        ```

1. Add a new file to the **scripts** folder named **auth.ts**

    1. Add the following to the **auth.ts** file. Note that there is a token named `[application-id-from-registration]` that must be replaced. Use the value of the Application Id copied from the Application Registration page.

    ```typescript
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
    ```

1. Locate the file **scripts/client.ts**
    1. Add the following line to the bottom of **scripts/client.ts**

      ```typescript
      export * from './auth';
      ```

1. Refresh the Tab in Microsoft Teams. Click the **Get MSGraph Data** button to invoke the authentication and call to graph.microsoft.com.


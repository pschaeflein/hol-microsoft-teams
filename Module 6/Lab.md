# Developing Extensions for Microsoft Teams - Module 6
----------------

This lab creates a Microsoft Teams app from the Tab and Bot created previously along with a connector.

### Office 365 Connector & Webhooks

In Microsoft Teams, full functionality for Office 365 Connectors is restricted to connectors that have been published to the Office Store. However, communicating with Microsfot Teams using Office 365 Connectors is identical to using the Incoming Webhook. This lab will show the messaging mechanics via the Webhook feature and then show the Teams user interface experience for registering a connector.

### Incoming Webhook

1. Click **Teams** in the left panel, then select a Team.
1. Select the **General** Channel in the selected team.
1. Click **...** next to the channel name, then select **Connectors**.

1. Select **Incoming Webhook** from the list, then click **Add**.
1. Enter a name for the webhook, upload an image to associate with the data from the webhook, then select **Create**.
1. Click the button next to the webhook URL to copy it.  (You will use the webhook URL in a subsequent step.)
1. Click **Done**.
1. Close the **Connectors** dialog.

### Create a simple Connector Card message to the webhook

1. Copy the `sample-connector-message.json` file from the `Lab Files` folder to your development machine.
1. Open a **PowerShell** window, go to the directory that contains the `sample-connector-message.json`, and enter the following commands:

    ```powershell
    $message = Get-Content .\sample-connector-message.json
    $url = <YOUR WEBHOOK URL>
    Invoke-RestMethod -ContentType="application/json" -Body $message -Uri <YOUR WEBHOOK URL> -Method Post
    ```

    > **Note:** Replace **&lt;YOUR WEBHOOK URL&gt;** with the webhook URL you saved when you created the **Incoming Webhook** connector.

1. When the POST succeeds, you will see a simple 1 outputted by the Invoke-RestMethod cmdlet.
1. Check the Conversations tab in the Microsoft Teams application. You will see the new  card message posted to the conversation.

    > Note: The action buttons will not work. Action buttons work only for Connectors registered and published in the Office Store.

### Run the ngrok secure tunnel application

1. Open a new **Command Prompt** window.
1. Change to the directory that contains the ngrok.exe application.
1. Run the command `ngrok http [port] -host-header=localhost:[port]` (Replace [port] with the port portion of the URL noted above.)
1. The ngrok application will fill the entire prompt window. Make note of the Forwarding address using https. This address is required in the next step.
1. Minimize the ngrok Command Prompt window. It is no longer referenced in this lab, but it must remain running.

### Office 365 Connector registration
The following steps are used to register an Office 365 Connector.

1. Register the Connector on the [Connectors Developer Dashboard](https://go.microsoft.com/fwlink/?LinkID=780623). Log on the the site and click **New Connector**.
1. On the **New Connector** page:

    1. Complete the Name and Description as appropriate for your connector.

    1. In the Events/Notifications section the list of events are displayed when registering the Connector in the Teams user inteface on a consent dialog. The Connector framework will only allow cards sent by your connector  to have **Actions URLs** that match what is provided here.

    1. The **Landing page for your users for Groups or Teams** is a URL that is rendered by the Microsoft Teams Application when users initiate the registration flow from a channel. This page is rendered in a popup provided by Teams. The **Redirect URLs** is a list of valid URLs to which the completed registration information can be sent. This functionality is similar to the Redirect URL processing for Azure Active Directory apps.

        For this lab, ensure that the hostname matches the ngrok forwarding address. For the landing page, append `/api/connector/landing` to the hostname. For the redirect page, append `/api/connector/redirect` to the hostname.

    1. In the **Enable this integration for** section, both **Group** and **Microsoft Teams** must be selected.
    1. Agree to the terms and conditions and click **Save**

1. The registration page will refresh with additional buttons in the integration section. The buttons provide sample code for the **Landing** page and a `manifest.json` file for a Teams app. Save both of these assets.

### Add Connector to existing Bot

In Visual Studio 2017, open the Bot solution from the previous modules. This bot will serve as the foundation for the complete Microsoft Teams app.

1. Open the `manifest.json` file in the solution's `Manifest` folder.
1. Replace the empty `connectors` node in the `manifest.json` file with the `connectors` node from the downloaded manifest. Save and close `manifest.json`
1. Open the file `WebApiConfig.cs` in the `App_Start` folder.
1. Modify the route configuration as shown. The original `routeTemplate` is `"api/{controller}/{id}"`. Replace the `id` token with the `action` token. Once complete, the statement should read as follows.

    ```cs
    config.Routes.MapHttpRoute(
      name: "DefaultApi",
      routeTemplate: "api/{controller}/{action}",
      defaults: new { id = RouteParameter.Optional }
    );
    ```

1. Right-click on the `Controllers` folder and select **Add | Controller...** Choose **Web API 2 Controller - Empty** and click **Add**. Name the new controller **ConnectorController** and click **Add**.
1. Add the following to the top of the `ConnectorController`.

    ```cs
    using System.Threading.Tasks;
    using System.Net.Http.Headers;
    ```

1. Add the following `Landing` method to the `ConnectorController`.

    ```cs
    [HttpGet]
    public async Task<HttpResponseMessage> Landing()
    {
      var htmlBody = "<html><title>Set up connector</title><body>";
      htmlBody += "<H2>Adding your Connector Portal-registered connector</H2>";
      htmlBody += "<p>Click the button to initiate the registration and consent flow for the connector in the selected channel.</p>";
      htmlBody += "<a href='https://outlook.office.com/connectors/Connect?state=myAppsState&app_id=ef5b13e5-261c-47f3-a7a8-6f00ef3b9930&callback_url=https://1ad3dcd5.ngrok.io/api/connector/redirect'>";
      htmlBody += "<img src='https://o365connectors.blob.core.windows.net/images/ConnectToO365Button.png' alt='Connect to Office 365'></img >";
      htmlBody += "</a>";

      var response = Request.CreateResponse(HttpStatusCode.OK);
      response.Content = new StringContent(htmlBody);
      response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
      return response;
    }
    ```

1. Add the following `Redirect` method to the `ConnectorController` class.

    ```cs
    [HttpGet]
    public async Task<HttpResponseMessage> Redirect()
    {
      // Parse register message from connector, find the group name and webhook url
      //var query = req.query;
      var query = Request.GetQueryNameValuePairs();
      string webhook_url = query.LastOrDefault(p => p.Key.Equals("webhook_url")).Value;
      var group_name = query.LastOrDefault(p => p.Key.Equals("group_name")).Value;
      var appType = query.LastOrDefault(p => p.Key.Equals("app_type")).Value;
      var state = query.LastOrDefault(p => p.Key.Equals("state")).Value;

      var htmlBody = "<html><body><H2>Registered Connector added</H2>";
      htmlBody += "<li><b>App Type:</b> " + appType + "</li>";
      htmlBody += "<li><b>Group Name:</b> " + group_name + "</li>";
      htmlBody += "<li><b>State:</b> " + state + "</li>";
      htmlBody += "<li><b>Web Hook URL stored:</b> " + webhook_url + "</li>";
      htmlBody += "</body></html>";

      var response = Request.CreateResponse(HttpStatusCode.OK);
      response.Content = new StringContent(htmlBody);
      response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
      return response;
    }
    ```

1. Press **F5** to run the application. This will also build the app package.
1. Re-sideload the application following the steps used earlier.

### Add Connector to a channel

1. Click **...** next to the channel name, then select **Connectors**.

1. Scroll to the bottom of the connector list. A section named **Sideloaded** contains the Connector described by the app. Click **Configure**.

1. An information dialog is shown with the general and notification information described on the Connector Developer portal. Click the **Visit site to install** button.

1. Click the **Connect to Office 365** button. Office 365 will process the registration flow, which may include login and Team/Channel selection. Make note of teh selected Teamd-Channel and click **Allow**.

1. The dialog will display the **Redirect** action which presents the information registration provided by Office 365. In a production application, this information must be presisted and used to sent notifications to the channel.

    > Note: Before your Connector can receive callbacks for actionable messages, you must register it and publish the app.

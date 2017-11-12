# Developing Extensions for Microsoft Teams - Module 3
----------------

This Module introduces the Bot Framework template and its capabilities for scaffolding a project and testing its functionality. In this exercise, you will create a basic Teams bot.

1. Launch Visual Studio 2017 as an administrator
1. In Visual Studio 2017, select **File | New | Project**
1. Create a new Visual C# project using the **Bot Application Template**

1. Build the solution to download all configured NuGet packages.

The Bot Application template is a fully functional Echo Bot that takes the user's text utterance as input and returns it as output. In order to run the bot inside Microsoft Teams:

- The bot must be accessible from the internet
- The bot must be registered with the Bot Connector
- The AppId and AppPassword from the Bot Framework registration page have to be recorded in the project's web.config
- The bot must be added to Microsoft Teams

Before registering the bot, note the URL configured for the solution in Visual Studio.

1. In Solution Explorer, double-click on **Properties**.
1. In the **Properties** designer, select the **Web** tab.
1. Note the **Project URL**.

### Run the ngrok secure tunnel application

1. Open a new **Command Prompt** window.
1. Change to the directory that contains the ngrok.exe application.
1. Run the command `ngrok http [port] -host-header=localhost:[port]` *(Replace [port] with the port portion of the URL noted above.)*
1. The ngrok application will fill the entire prompt window. Make note of the Forwarding address using https. This address is required in the next step.
1. Minimize the ngrok Command Prompt window. It is no longer referenced in this lab, but it must remain running.

### Register the bot

1. Go to the Microsoft Bot Framework portal at https://dev.botframework.com and sign in. (The bot registration portal accepts a Work or School Account or a Microsoft Account.)
1. Click Register. (If the Register button is not shown, click **My bots** in the top navigation.)
1. Complete the Bot profile section, entering a Display name, Bot handle and description.

1. Complete the Configuration section.
    1. For the Messaging endpoint, use the Forwarding https address from ngrok prepended to the route to the MessagesController in the Visual Studio project.
    1. Click the **Create Microsoft App ID and password button**. This opens a new browser tab/window.
    1. In the new browser tab/window the application is registered in Azure Active Directory. Click **Generate an app password to continue**.
    1. An app password is generated. Copy the password and save it. You will use it in a subsequent step.
    1. Click **OK**. This closes the popup.
    1. Click the **Finish and go back to Bot Framework** button. This closes the new browser tab/window and populates the app Id in the **Paste your app ID below to continue textbox**.

1. Scroll to the bottom of the page. Agree to the Privacy statement, Terms of use, and Code of conduct and click the **Register** button. Once the Bot is created, click **OK** to dismiss the pop-up.

The **Connect to channels** page is displayed for the newly-created bot. The bot must be connected to Microsoft Teams.

1. Click the **Teams** logo.

1. Once the connection is complete, ensure the connection is Enabled and click **Done**

The bot registration is complete.

> Clicking on Settings in the top navigation will re-display the profile and configuration sections. This can be used to update the Messaging endpoint in the event ngrok is stopped, or the bot is moved to staging/production.

### Configure the web project

The bot project must be configured with information from the registration.

1. In Visual Studio, open the **Web.config** file. Locate the `<appSettings>` section.
1. Enter the `BotId` value. the `BotId` is the **Bot handle** from the **Configuration** section of the registration.
1. Enter the `MicrosoftAppId`. The `MicrosoftAppId` is the app ID from the **Configuration** section of the registration.
1. Enter the `MicrosoftAppPassword`. The `MicrosoftAppPassword` is the auto-generated app password displayed in the pop-up during registration.

    > If you do not have the app password, the bot must be deleted and re-registered. An app password cannot be reset nor displayed.

### Test the bot using the portal

The Bot registration portal can be used to test the bot.

1. Ensure ngrok is still running, and the Messaging endpoint of the bot registration is using the hostname shown as the Forwarding https address in ngrok.
1. In Visual Studio, press **F5** to start the project.
1. Once the **default.htm** page is displayed, return to the [Bot registration portal](https://dev.botframework.com/bots).
1. Select your bot.
1. In the top-right corner of the page, click the **Test** button.
1. Enter a message and press **Enter**. The message is echoed back along with the length of the mesage.

    > If the message cannot be sent, there is an error in the configuration of the Bot registration, ngrok and Visual Studio. The request should be visible in the ngrok command window. For additional detail on the request in ngrok, open the address `http://localhost:4040`. If no requests are displayed in ngrok, then the Messaging endpoint has the wrong hostname, or there is a disruption in the network connectivity.

### Configure Visual Studio to Package bot

Packaging a bot for Microsoft Teams is identical to packaging a tab. A manifest file (and related resources) are compressed into a zip file and added to a team.

Perform the following in Visual Studio. (Stop debugging before continuing. Leave ngrok running.)
1.  Right-click on the project, choose **Add | New Folder**. Name the folder **Manifest**.
1. Add the displayed files from the **Lab Files** folder of this repository.

1. Open the **manifest.json** file just added to the project.

    The **manifest.json** file requires several updates:
    - The `id` property must contain the app ID from registration. Replace the token `[microsoft-app-id]` with the app ID.
    - The `packageName` property must contain a unique identifier. The convention is to use the bot's URL in reverse format. Replace the token `[from-ngrok]` with the unique identifier from the Forwarding address.
    - Similarly, the `developer` property has three URLs that should match the hostname of the Messaging endpoint. Replace the token `[from-ngrok]` with the unique identifier from the Forwarding address.
    - The `botId` property (in the `bots` collection property) also requires the app ID from registration. Replace the token `[microsoft-app-id]` with the app ID.
    - Save and close the **manifest.json** file.

1. Update the Visual Studio project to compress the Manifest folder during build.
    - In Solution Explorer, right-click on the project and choose **Unload Project**. If prompted, click **Yes** to save changes.

    - Right-click on the project file and choose **Edit [project-name].csproj**. *(In the image, the project name is teams-bot1.)*

    - Scroll to the bottom of the file. Add the following Target to the file. *(Be sure to add the target outside of the comment.)* This target will invoke a custom build task to compress the files in the Manfest directory.

      ```xml
      <Target Name="AfterBuild">
        <ZipDir InputBaseDirectory="manifest"
                OutputFileName="$(OutputPath)\$(MSBuildProjectName).zip"
                OverwriteExistingFile="true"
                IncludeBaseDirectory="false" />
      </Target>
      ```

    - Add the following Task element to the **.csproj** file.

      ```xml
      <UsingTask TaskName="ZipDir" TaskFactory="CodeTaskFactory"
                AssemblyFile="$(MSBuildToolsPath)\Microsoft.Build.Tasks.v4.0.dll">
        <ParameterGroup>
          <InputBaseDirectory ParameterType="System.String" Required="true" />
          <OutputFileName ParameterType="System.String" Required="true" />
          <OverwriteExistingFile ParameterType="System.Boolean" Required="false" />
          <IncludeBaseDirectory ParameterType="System.Boolean" Required="false" />
        </ParameterGroup>
        <Task>
          <Reference Include="System.IO.Compression" />
          <Reference Include="System.IO.Compression.FileSystem" />
          <Using Namespace="System.IO.Compression" />
          <Code Type="Fragment" Language="cs"><![CDATA[
            if (File.Exists(OutputFileName))
            {
              if (!OverwriteExistingFile)
              {
                return false;
              }
              File.Delete(OutputFileName);
            }
            ZipFile.CreateFromDirectory
            (
              InputBaseDirectory, OutputFileName,
              CompressionLevel.Optimal, IncludeBaseDirectory
            );
          ]]></Code>
        </Task>
      </UsingTask>
      ```

    - Save and close the project file.
    - In **Solution Explorer**, right-click on the project and choose **Reload Project**.

1. Press **F5** to run the project. The new AfterBuild target will run, creating a zip file in the build output folder (bin\\)

### Sideload app into Microsoft Teams

Although not strictly necessary, in this lab the bot will be added to a new Team.

1. In the Microsoft Teams application, click the **Add team** link. Then click the **Create team** button.

1. Enter a team name and description. In this example, the Team is named **teams-bot-1**. Click **Next**.
1. Optionally, invite others from your organization to the team. This step can be skipped in this lab.
1. The new team is shown. In the left-side panel, click the ellipses next to the team name. Choose **View team** from the context menu.

1. On the View team display, click **Apps** in the tab strip. Then click the **Sideload an app** link at the bottom right corner of the application.
1. Select the **teams-bot1.zip** file from the *bin* folder. Click **Open**.
1. The app is displayed. Notice information about the app from the manifest (Description and Icon) is displayed.

The app is now sideloaded into the Microsoft Teams application and the bot is available.

### Interact with the Bot

In the General channel for the team, a message is created announcing the bot has been added to the Team. To interact with the bot, @ mention the bot.

As you test the bot, you will notice that the character count is not correct. You can set breakpoints in the Visual Studio project to debug the code. (Remember, however, that the count was correct using the registration portal.) Later modules of this training will review how to remove mentions from the message.

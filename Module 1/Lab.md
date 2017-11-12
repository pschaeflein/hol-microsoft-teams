# Developing Extensions for Microsoft Teams - Module 1
----------------
In this module lab, you will learn the steps to generate, package and test your Microsoft Teams application.

## Create and Test a Basic Teams App using Yeoman

This exercise introduces the Yeoman generator and its capabilities for scaffolding a project and testing its functionality.
In this exercise, you will create a basic Teams app.

1. Open a **Command Prompt** window.
1. Change to the directory where you will create the tab.

  > **Note:** Directory paths can become quite long after node modules are imported.  **We suggest you use a directory name without spaces in it and create it in the root folder of your drive.**  This will make working with the solution easier in the future and protect you from potential issues associated with long file paths. In this example, we use `c:\Dev` as the working directory.

1. Type `md teams-mod1` and press **Enter**.
1. Type `cd teams-mod1` and press **Enter**.

### Run the Yeoman Teams Generator

1. Type `yo teams` and press **Enter**.

1. When prompted:
    1. Accept the default **teams-mod-1** as your solution name and press **Enter**.
    1. Select **Use the current folder** for where to place the files and press **Enter**.
1. The next set of prompts asks for specific information about your Teams App:
    1. Accept the default **teams mod1** as the name of your Microsoft Teams App project and press **Enter**.
    1. Enter your name and press **Enter**.
    1. Enter **https://tbd.ngrok.io** as the URL where you will host this tab and press **Enter**. (We will change this URL later.)
    1. Accept the default selection of **Tab** for what you want to add to your project and press **Enter**.
    1. Enter **Module 1** as the default tab name and press **Enter**.

  >**Note:** At this point, Yeoman will install the required dependencies and scaffold the solution files along with the basic tab. This might take a few minutes.

### Run the ngrok secure tunnel application

1. Open a new **Command Prompt** window.
1. Change to the directory that contains the ngrok.exe application.
1. run the command `ngrok http 3007`
1. The ngrok application will fill the entire prompt window. Make note of the Forwarding address using https. This address is required in the next step.
1. Minimize the ngrok Command Prompt window. It is no longer referenced in this exercise, but it must remain running.

### Update the Teams app manifest and create package

When the solution was generated, we used a placeholder URL. Now that the tunnel is running, we need to use the actual URL that is routed to our computer.

1. Return to the first **Command Prompt** window in which the generator was run.
1. Launch Visual Studio Code by running the command `code .`
1. Open the **manifest.json** file in the **manifest** folder.
1. Replace all instances of **tbd.ngrok.io** with the HTTPS Forwarding address from the ngrok window. There are 6 URLs that need to be changed.
1. Save the **manifest.json** file.
1. In the **Command Prompt** window, run the command `gulp manifest`. This command will create the package as a zip file in the **package** folder
1. Build the webpack and start the Express web server by running the following commands:

    ```shell
    gulp build
    gulp serve
    ```

    > Note: The gulp serve process must be running in order to see the tab in the Teams application. When the process is no longer needed, press `CTRL+C` to cancel the server.

### Sideload app into Microsoft Teams

1. In the Microsoft Teams application, click the **Add team** link. Then click the **Create team** button.
1. Enter a team name and description. Click **Next**.
1. Optionally, invite others from your organization to the team. This step can be skipped in this lab.
1. The new team is shown. In the left-side panel, click the elipses next to the team name. Choose **View team** from the context menu.
1. On the View team display, click **Apps** in the tab strip. Then click the **Sideload an app** link at the bottom right corner of the application.
1. Select the **teams-mod-1.zip** file from the **package** folder. Click **Open**.
1. The app is displayed. Notice information about the app from the manifest (Description and Icon) is displayed.

The app is now sideloaded into the Microsoft Teams application and the Tab is available in the **Tab Gallery**.

### Add Tab to Team view

1. Tabs are not automatically displayed for the Team. To add the tab, click on the **General** channel in the Team.
1. Click the **+** icon at the end of the tab strip.
1. In the Tab gallery, sideloaded tabs are displayed in the **Tabs for your team** section. Tabs in this section are arranged alphabetically. Select the tab created in this lab.
1. The generator creates a configurable tab. When the Tab is added to the Team, the configuration page is displayed. Enter any value in the **Setting** box and click **Save**.
1. The value entered will then be displayed in the Tab window.
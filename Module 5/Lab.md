# Developing Extensions for Microsoft Teams - Module 5
----------------
This lab extends the bot from Module 4 with Microsoft Teams functionality called Compose Extension. Compose Extensions provide help for users when composing a message for posting in a channel or in 1:1 chats.

1. Open the `MessagesController.cs` file in the `Controllers` folder.
1. Locate the `Post` method. Replace the method the following snippet. Rather than repeating if statements, the logic is converted to a switch statement. Compose Extensions are posted to the bot via an `Invoke` message.

    ```cs
    public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
    {
      switch (activity.Type)
      {
        case ActivityTypes.Message:
          await Conversation.SendAsync(activity, () => new Dialogs.RootDialog());
          break;

        case ActivityTypes.ConversationUpdate:
          await HandleSystemMessageAsync(activity);
          break;

        case ActivityTypes.Invoke:
          var composeResponse = await ComposeHelpers.HandleInvoke(activity);
          var stringContent = new StringContent(composeResponse);
          HttpResponseMessage httpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK);
          httpResponseMessage.Content = stringContent;
          return httpResponseMessage;
          break;

        default:
          break;
      }
      var response = Request.CreateResponse(HttpStatusCode.OK);
      return response;
    }
    ```

1. In **Solution Explorer**, add a new class to the project. Name the class `BotChannelsData`. Replace the generated class with the code from file `Lab Files/BotChannelData.cs`.

    ```cs
    using System.Collections.Generic;

    namespace change_this
    {
      public class BotChannel
      {
        public string Title { get; set; }
        public string LogoUrl { get; set; }
      }

      public class BotChannels
      {
        public static List<BotChannel> GetBotChannels()
        {
          var data = new List<BotChannel>();
          data.Add(new BotChannel { Title = "Bing", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/bing.png" });
          data.Add(new BotChannel { Title = "Cortana", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/cortana.png" });
          data.Add(new BotChannel { Title = "Direct Line", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/directline.png" });
          data.Add(new BotChannel { Title = "Email", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/email.png" });
          data.Add(new BotChannel { Title = "Facebook Messenger", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/facebook.png" });
          data.Add(new BotChannel { Title = "GroupMe", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/groupme.png" });
          data.Add(new BotChannel { Title = "Kik", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/kik.png" });
          data.Add(new BotChannel { Title = "Microsoft Teams", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/msteams.png" });
          data.Add(new BotChannel { Title = "Skype", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/skype.png" });
          data.Add(new BotChannel { Title = "Skype for Business", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/skypeforbusiness.png" });
          data.Add(new BotChannel { Title = "Slack", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/slack.png" });
          data.Add(new BotChannel { Title = "Telegram", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/telegram.png" });
          data.Add(new BotChannel { Title = "Twilio (SMS)", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/sms.png" });
          data.Add(new BotChannel { Title = "Web Chat", LogoUrl = "https://dev.botframework.com/client/images/channels/icons/webchat.png" });
          return data;
        }
      }
    }
    ```

1. In **Solution Explorer**, add a new class to the project. Name the class `ComposeHelpers`. Add the code from the `Lab Files/ComposeHelpers.cs` file.

    ```cs
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Teams;
    using Microsoft.Bot.Connector.Teams.Models;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading.Tasks;

    namespace change_this
    {
      public class ComposeHelpers
      {
        public static async Task<HttpResponseMessage> HandleInvoke(Activity activity)
        {
          // these are the values specified in manifest.json
          string COMMANDID = "searchCmd";
          string PARAMNAME = "searchText";

          var unrecognizedResponse = new HttpResponseMessage(HttpStatusCode.BadRequest);
          unrecognizedResponse.Content = new StringContent("Invoke request was not recognized.");

          if (!activity.IsComposeExtensionQuery())
          {
            return unrecognizedResponse;
          }

          // This helper method gets the query as an object.
          var query = activity.GetComposeExtensionQueryData();
          if (query.CommandId == null || query.Parameters == null)
          {
            return unrecognizedResponse;
          }

          if (query.CommandId != COMMANDID)
          {
            return unrecognizedResponse;
          }

          var param = query.Parameters.FirstOrDefault(p => p.Name.Equals(PARAMNAME)).Value.ToString();
          if (String.IsNullOrEmpty(param))
          {
            return unrecognizedResponse;
          }

          // This is the response object that will get sent back to the compose extension request.
          ComposeExtensionResponse invokeResponse = null;

          // search our data
          var resultData = BotChannels.GetBotChannels().FindAll(t => t.Title.Contains(param));

          // format the results
          var results = new ComposeExtensionResult()
          {
            AttachmentLayout = "list",
            Type = "result",
            Attachments = new List<ComposeExtensionAttachment>(),
          };

          foreach (var resultDataItem in resultData)
          {
            var card = new ThumbnailCard()
            {
              Title = resultDataItem.Title,
              Images = new List<CardImage>() { new CardImage() { Url = resultDataItem.LogoUrl } }
            };

            var composeExtensionAttachment = card.ToAttachment().ToComposeExtensionAttachment();
            results.Attachments.Add(composeExtensionAttachment);
          }

          invokeResponse.ComposeExtension = results;

          // Return the response
          StringContent stringContent;
          try
          {
            stringContent = new StringContent(JsonConvert.SerializeObject(invokeResponse));
          }
          catch (Exception ex)
          {
            stringContent = new StringContent(ex.ToString());
          }
          var response = new HttpResponseMessage(HttpStatusCode.OK);
          response.Content = stringContent;
          return response;
        }

      }
    }
    ```

1. Open the `manifest.json` file in the `Manifest` folder. Locate the `composeExtensions` node and replace it with the following snippet. Replace the `[MicrosoftAppId]` token with the app ID from the settings page of the bot registration (https://dev.botframework.com).

    ```json
    "composeExtensions": [
      {
        "botId": "[MicrosoftAppId]",
        "scopes": [
          "team"
        ],
        "canUpdateConfiguration": true,
        "commands": [
          {
            "id": "searchCmd",
            "description": "Search Bot Channels",
            "title": "Bot Channels",
            "initialRun": false,
            "parameters": [
              {
                "name": "searchText",
                "description": "Enter your search text",
                "title": "Search Text"
              }
            ]
          }
        ]
      }
    ],
    ```

1. Press **F5** to re-build the app package and start the debugger.
1. Re-sideload the app. Since the `manifest.json` has been updated, the bot must be re-sideloaded to the Microsoft Teams application.

### Invoke the Compose Extension

The Compose Extension is configured for use in a Channel (due to the scopes entered in the manifest.) The extension is invoked by clicking the elipsis below the compose box and selecting the bot.

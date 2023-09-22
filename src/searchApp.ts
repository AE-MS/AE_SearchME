import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  TaskModuleRequest,
  TaskModuleResponse,
  MessagingExtensionAction,
  MessagingExtensionActionResponse
} from "botbuilder";

const urlDialogTriggerValue = 500;
const cardDialogTriggerValue = 501;

const adaptiveCardBotJson = {
  "contentType": "application/vnd.microsoft.card.adaptive",
  "content": {
      "type": "AdaptiveCard",
      "body": [
          {
              "type": "TextBlock",
              "text": "Here is a ninja cat:"
          },
          {
              "type": "Image",
              "url": "http://adaptivecards.io/content/cats/1.png",
              "size": "Medium"
          }
      ],
      "actions": [
          {
              "type": "Action.Submit",
              "title": "Submit"
          }
      ],
      "version": "1.0"
  }
}

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const adaptiveCard = CardFactory.heroCard(
        `${obj.package.name}`,
        `${obj.package.description}`,
        null, // No images
      [{
          type: 'invoke',
          title: "Show URL Task Module",
          value: {
              type: 'task/fetch',
              data: urlDialogTriggerValue
          }
        },
        {
          type: 'invoke',
          title: "Show Adaptive Card Task Module",
          value: {
              type: 'task/fetch',
              data: cardDialogTriggerValue
          }
        }]
      );

      const preview = CardFactory.heroCard(`${obj.package.name}???`);
      const attachment = { ...adaptiveCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    const cardTaskFetchValue = taskModuleRequest.data.data;
    console.log(`FETCH VALUE: ${cardTaskFetchValue}`);

    var taskInfo = {};

    switch (cardTaskFetchValue) {
      case urlDialogTriggerValue:
        taskInfo = {
          url: "https://m365playgrcb3596tab.z5.web.core.windows.net/index.html#/tab",
          fallbackUrl: "https://m365playgrcb3596tab.z5.web.core.windows.net/index.html#/tab",
          height: 510,
          width: 450,
          title: "URL Dialog",
        };
        break;

      case cardDialogTriggerValue:
        taskInfo = {
          card: adaptiveCardBotJson,
          height: 510,
          width: 450,
          title: "Adaptive Card Dialog",
        };
        break;
    }

    return Promise.resolve({
      task: {
        type: 'continue',
        value: taskInfo,
      }
    });
  }

  override handleTeamsTaskModuleSubmit(_context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`HANDLING DIALOG SUBMIT: ${JSON.stringify(taskModuleRequest)}`);

    return Promise.resolve({
      task: {
        type: 'continue',
        value: {
          url: "https://m365playgrcb3596tab.z5.web.core.windows.net/index.html#/tab",
          fallbackUrl: "https://m365playgrcb3596tab.z5.web.core.windows.net/index.html#/tab",
          height: 510,
          width: 450,
          title: "URL Dialog",
        }
      }
    });

    // return Promise.resolve({
    //   task: {
    //       type: 'message',
    //       value: 'Thanks!'
    //   }
    // });
  }

  override handleTeamsMessagingExtensionSubmitAction(_context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    console.log(`HANDLING ME SUBMIT ACTION: ${JSON.stringify(action)}`);
    
    const data = action.data;
    const heroCard = CardFactory.heroCard(data.title, data.text);
    heroCard.content.subtitle = data.subTitle;
    const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };

    return Promise.resolve({
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [],
      },
    });

  }

}

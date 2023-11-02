import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  ActionTypes,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  TaskModuleRequest,
  TaskModuleResponse,
  MessagingExtensionAction,
  MessagingExtensionActionResponse
} from "botbuilder";

const urlDialogTriggerValue = "requestUrl";
const cardDialogTriggerValue = "requestCard";
const messagePageTriggerValue = "requestMessage";
const noResponseTriggerValue = "requestNoResponse";

const pageDomain = "localhost:53000";
// const pageDomain = "helloworld36cffe.z5.web.core.windows.net";

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
              "data": { data: urlDialogTriggerValue },
              "type": "Action.Submit",
              "title": "Request URL Dialog"
          },
          {
            "data": { data: cardDialogTriggerValue },
            "type": "Action.Submit",
            "title": "Request Card Dialog"
          },
          {
            "data": { data: messagePageTriggerValue },
            "type": "Action.Submit",
            "title": "Request Message"
          },
          {
            "data": { data: noResponseTriggerValue },
            "type": "Action.Submit",
            "title": "Request No Response (close Dialog)"
          },
      ],
      "version": "1.0"
  }
}

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Called when a search query is triggered by the user from the query box.
  // This function is called in response to the "composeExtension/query" invoke message
  // Possible repsonses: https://learn.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/search-commands/respond-to-search?tabs=typescript#respond-to-user-requests
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;
    console.log(`SEARCH QUERY: ${JSON.stringify(query)}`);
    if (searchQuery === "config") {
      return {
        composeExtension: {
          type: 'config',
          suggestedActions: {
              actions: [
                  {
                    title: "Config Action Title",
                    type: ActionTypes.OpenUrl,
                    value: `https://${pageDomain}/index.html?page=config#/tab`
                  },
              ],
          },
        },
      };
    } else if (searchQuery === "auth") {
      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
              actions: [
                  {
                    title: "Auth Action Title",
                    type: ActionTypes.OpenUrl,
                    value: `https://${pageDomain}/index.html?page=auth#/tab`
                  },
              ],
          },
        },
      };
    } else if (searchQuery === "sso") {
      return {
        composeExtension: {
          type: 'silentAuth',
          suggestedActions: {
              actions: [
                  {
                    title: "SSO?",
                    type: ActionTypes.OpenUrl,
                    value: `https://${pageDomain}/index.html?page=auth#/tab`
                  },
              ],
          },
        },
      };
    } else if (searchQuery === "message") {
      return {
        composeExtension: {
          type: 'message',
          text: "This is a message",
        },
      };
    } else {
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
  }

  private getRandomIntegerBetween(min: number, max: number): number {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  private createUrlTaskModuleResponse(): Promise<TaskModuleResponse> {
    return Promise.resolve({
      task: {
        type: 'continue',
        value: {
          url: `https://${pageDomain}/index.html?randomNumber=${this.getRandomIntegerBetween(1, 1000)}#/tab`,
          fallbackUrl: "https://thisisignored.example.com/",
          height: 510,
          width: 450,
          title: "URL Dialog",
        }
      }
    });
  }

  private createCardTaskModuleResponse(): Promise<TaskModuleResponse> {
    return Promise.resolve({
      task: {
        type: 'continue',
        value: {
          card: adaptiveCardBotJson,
          height: 510,
          width: 450,
          title: "Adaptive Card Dialog",
        }
      }
    });
  }

  private createMessageTaskModuleResponse(): Promise<TaskModuleResponse> {
    return Promise.resolve({
      task: {
          type: 'message',
          value: `Hello! This is a message!`,
      }
    });
  }
  
  override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`TASK MODULE FETCH. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return this.createResponseToTaskModuleRequest(taskModuleRequest);
  }

  override handleTeamsTaskModuleSubmit(_context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`HANDLING DIALOG SUBMIT. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return this.createResponseToTaskModuleRequest(taskModuleRequest);
  }

  override handleTeamsConfigSubmit(_context: TurnContext, configData: any): Promise<any> {
    console.log(`HANDLING CONFIG SUBMIT, configData: ${configData}`);

    return;
  }

  override handleTeamsConfigFetch(_context: TurnContext, configData: any): Promise<any> {
    console.log(`HANDLING CONFIG FETCH, configData: ${configData}`);

    return;
  }

  private createResponseToTaskModuleRequest(taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    const taskRequestData = taskModuleRequest.data.data;

    switch (taskRequestData) {
      case urlDialogTriggerValue:
        return this.createUrlTaskModuleResponse();

      case cardDialogTriggerValue:
        return this.createCardTaskModuleResponse();

      case messagePageTriggerValue:
        return this.createMessageTaskModuleResponse();

      case noResponseTriggerValue:
        return;

      default:
        return Promise.resolve({
          task: {
              type: 'message',
              value: `The submitted data did not contain a valid request (submitted data: ${taskModuleRequest.data})`,
          }
      });
    }
  }  

  // I have not been able to get this function ever called
  override async handleTeamsMessagingExtensionConfigurationSetting(_context: TurnContext, settings: any): Promise<void> {
    console.log(`CONFIG WAS SET. Settings: ${JSON.stringify(settings)}`);
  }  

  // I have no idea when this function can ever be called
  override async handleTeamsMessagingExtensionConfigurationQuerySettingUrl(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResponse> {
    console.log(`CONFIG QUERY SETTING URL. Query: ${JSON.stringify(query)}, context: ${JSON.stringify(context)}`);

    return {
        composeExtension: {
            type: 'config',
            suggestedActions: {
                actions: [
                    {
                      title: "The title",
                      type: ActionTypes.OpenUrl,
                      value: `https://${pageDomain}/index.html#/tab`
                    },
                ],
            },
        },
    };
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

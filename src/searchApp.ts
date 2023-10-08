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
              "data": "requestUrl",
              "type": "Action.Submit",
              "title": "Request URL Dialog"
          },
          {
            "data": "requestCard",
            "type": "Action.Submit",
            "title": "Request Card Dialog"
          },
          {
            "data": "requestMessage",
            "type": "Action.Submit",
            "title": "Request Message"
          },
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

  private createUrlTaskModuleResponse(): Promise<TaskModuleResponse> {
    return Promise.resolve({
      task: {
        type: 'continue',
        value: {
          url: "https://helloworld36cffe.z5.web.core.windows.net/index.html#/tab",
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

  override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    const cardTaskFetchValue = taskModuleRequest.data.data;
    console.log(`FETCH VALUE: ${cardTaskFetchValue}`);

    var taskInfo = {};

    switch (cardTaskFetchValue) {
      case urlDialogTriggerValue:
        return this.createUrlTaskModuleResponse();

      case cardDialogTriggerValue:
        return this.createCardTaskModuleResponse();
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

    if (taskModuleRequest.data === "requestUrl") {
      return this.createUrlTaskModuleResponse();
    } else if (taskModuleRequest.data === "requestCard") {
      return this.createCardTaskModuleResponse();      
    } else if (taskModuleRequest.data === "requestMessage") {
      return Promise.resolve({
        task: {
            type: 'message',
            value: `Hello! This is a message!`,
        }
      });        
    } else if (taskModuleRequest.data === "requestConfig") {
      return Promise.resolve({
        task: {
            type: 'message',
            value: `The submitted data did not contain a valid request (submitted data: ${taskModuleRequest.data})`,
        }
      });          
    } else if (taskModuleRequest.data === "requestNoResponse") {
      return;
    }
    else {
        return Promise.resolve({
            task: {
                type: 'message',
                value: `The submitted data did not contain a valid request (submitted data: ${taskModuleRequest.data})`,
            }
        });
    }
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

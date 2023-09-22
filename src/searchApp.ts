import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  TaskModuleRequest,
  TaskModuleResponse
} from "botbuilder";

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
          title: "Show Task Module",
          value: {
              type: 'task/fetch',
              data: 500
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
    var taskInfo = {
      url: "https://m365playgrcb3596tab.z5.web.core.windows.net/index.html#/tab",
      fallbackUrl: "https://m365playgrcb3596tab.z5.web.core.windows.net/index.html#/tab",
      height: 510,
      width: 450,
      title: "A Task Module",
    };

    return Promise.resolve({
      task: {
        type: 'continue',
        value: taskInfo,
      }
    });
  }  

}

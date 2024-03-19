const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const helloWorldCard = require("./adaptiveCards/helloWorldCard.json");

const { DiceRoll } = require('@dice-roller/rpg-dice-roller');

class ActionApp extends TeamsActivityHandler {
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const template = new ACData.Template(helloWorldCard);

    const roll = new DiceRoll(action.data.text ?? "");
    console.log(roll.output);

    const card = template.expand({
      $root: {
        title: action.data.title ?? "",
        subTitle: action.data.subTitle ?? "",
        text: roll.output,
      },
    });
    const attachment = CardFactory.adaptiveCard(card);
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
      },
    };
  }
}
module.exports.ActionApp = ActionApp;

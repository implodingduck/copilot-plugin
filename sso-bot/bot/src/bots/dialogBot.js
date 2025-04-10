// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, AttachmentLayoutTypes } = require('botbuilder');


const chartpayload1 = {
    "type": "AdaptiveCard",
    "version": "1.5",
    "body": [
        {
            "type": "TextBlock",
            "text": "Basic",
            "size": "extraLarge"
        },
        {
            "type": "Chart.Gauge",
            "value": 50,
            "segments": [
                {
                    "legend": "Low risk",
                    "size": 33,
                    "color": "good"
                },
                {
                    "legend": "Medium risk",
                    "size": 34,
                    "color": "warning"
                },
                {
                    "legend": "High risk",
                    "size": 33,
                    "color": "attention"
                }
            ]
        }
    ]
}




/**
 * DialogBot class extends TeamsActivityHandler to handle Teams activities.
 */
class DialogBot extends TeamsActivityHandler {
    /**
     * Creates an instance of DialogBot.
     * @param {ConversationState} conversationState - The state management object for conversation state.
     * @param {UserState} userState - The state management object for user state.
     * @param {Dialog} dialog - The dialog to be run by the bot.
     */
    constructor(conversationState, userState, dialog) {
        super();

        if (!conversationState) {
            throw new Error('[DialogBot]: Missing parameter. conversationState is required');
        }
        if (!userState) {
            throw new Error('[DialogBot]: Missing parameter. userState is required');
        }
        if (!dialog) {
            throw new Error('[DialogBot]: Missing parameter. dialog is required');
        }

        this.conversationState = conversationState;
        this.userState = userState;
        this.dialog = dialog;
        this.dialogState = this.conversationState.createProperty('DialogState');
        this.randomnumber = Math.floor(Math.random() * 100);

        this.onMessage(this.handleMessage.bind(this));
    }

    randompayload() {
        const payload = {
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": "Single value",
                        "size": "extraLarge",
                        "spacing": "large",
                        "separator": true
                    },
                    {
                        "type": "Chart.Gauge",
                        "value": this.randomnumber,
                        "valueFormat": "fraction",
                        "segments": [
                            {
                                "legend": "Used",
                                "size": this.randomnumber
                            },
                            {
                                "legend": "Unused",
                                "size": (100-this.randomnumber),
                                "color": "neutral"
                            }
                        ]
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Execute",
                        "title": "Randomize",
                        "data": {
                            "currentvalue": this.randomnumber
                        }
                    }
                ]
        };
        return payload;
    }

    /**
     * Handles incoming message activities.
     * @param {TurnContext} context - The context object for the turn.
     * @param {Function} next - The next middleware function in the pipeline.
     */
    async handleMessage(context, next) {
        console.log('Running dialog with Message Activity.');
        console.log(`Context: ${JSON.stringify(context)}`)
        // Run the Dialog with the new message Activity.
        console.log(`Message: ${context.activity.text}`)


        if (context.activity.text === 'login') {
            await this.dialog.run(context, this.dialogState);
        } else if (context.activity.text === 'card') {
            await context.sendActivity({
                attachments: [
                    CardFactory.adaptiveCard(chartpayload1),
                    CardFactory.adaptiveCard(this.randompayload()),
                ],
                attachmentLayout: AttachmentLayoutTypes.Carousel
            });
        } else if (context.activity.text === 'random') {
            await context.sendActivity(`This current random number is: ${this.randomnumber}`);
        }
        else {
            await context.sendActivity(`Please type "login" to start the authentication process... echoing back your message: ${context.activity.text}`);
        }


        await next();
    }

    async onInvokeActivity(context) {
        console.log(`Context: ${JSON.stringify(context)}`);
        console.log(`Invoke: ${context.activity.name}`);
        return super.onInvokeActivity(context);
    }

    async onAdaptiveCardInvoke(context, invokeValue) {
        console.log(`Context:`);
        console.log(`${JSON.stringify(context, null, 2)}`);
        console.log(`Invoke Value: ${invokeValue}`);
        this.randomnumber = Math.floor(Math.random() * 100);
        const payload = this.randompayload();
        const cardRes = {
            statusCode: StatusCodes.OK,
            type: 'application/vnd.microsoft.card.adaptive',
            value: payload
        };
        return {
            statusCode: 200,
            body: cardRes
        }
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     * @param {TurnContext} context - The context object for the turn.
     */
    async run(context) {
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);
    }
}

module.exports.DialogBot = DialogBot;
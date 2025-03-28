// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, DialogSet, DialogTurnStatus, OAuthPrompt, WaterfallDialog } = require('botbuilder-dialogs');
const { LogoutDialog } = require('./logoutDialog');
const { SimpleGraphClient } = require('../simpleGraphClient');
const { CardFactory } = require('botbuilder-core');

const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

/**
 * MainDialog class extends LogoutDialog to handle the main dialog flow.
 */
class MainDialog extends LogoutDialog {
    /**
     * Creates an instance of MainDialog.
     * @param {string} connectionName - The connection name for the OAuth provider.
     */
    constructor() {
        super(MAIN_DIALOG, process.env.connectionName);

        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
            connectionName: process.env.connectionName,
            text: 'Please Sign In',
            title: 'Sign In',
            timeout: 300000
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.loginStep.bind(this),
            this.ensureOAuth.bind(this),
            this.displayToken.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {TurnContext} context - The context object for the turn.
     * @param {StatePropertyAccessor} accessor - The state property accessor for the dialog state.
     */
    async run(context, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    /**
     * Prompts the user to sign in.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async promptStep(stepContext) {
        return await stepContext.beginDialog(OAUTH_PROMPT);
    }

    /**
     * Handles the login step.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async loginStep(stepContext) {
        const tokenResponse = stepContext.result;
        console.log(`Token response: ${JSON.stringify(tokenResponse)}`);
        if (!tokenResponse || !tokenResponse.token) {
            await stepContext.context.sendActivity('Login was not successful, please try again.');
            return await stepContext.endDialog();
        } else {
            try {
                const client = new SimpleGraphClient(tokenResponse.token);
                const me = await client.getMe();
                const title = me ? me.jobTitle : 'Unknown';
                await stepContext.context.sendActivity(`You're logged in as ${me.displayName} (${me.userPrincipalName}); your job title is: ${title}; your photo is: `);
                let photoBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAB4AAAAeCAIAAAC0Ujn1AAABeklEQVRIDe2UQcuDMAyG92+97OLBwwQRxpRNdAw9FCZsiohDCgqKiOB/20FoY7TV+eHtgx6SEJ/GN2kOimbtdA47cRXN+kcjbZcF0e0gTGjV9sNJaXW5PxFl1l1AX+5PBoVGmNDj6TZLZEEZWjVciEO2H6aMMmvI0OT9QTjkqoY7Cx2CMnRRd4iFXI/EW9DH0w2Bpm6Y0C1oudDDNVFWbEErmjUtE0XknZRpndIKsZArH3AZ2nQIYkE3LxuJGsvrSTJ/uh38Ca1olh+msNiq7fOyWeQKq9bOD4/E5P0ZHoVquFf/5ZHYI7HpkKFY3Q6irPBILLoGa206JC8bVmZRd1f/hX5cNVwkVEor7fxAaSO0qG952fhhOlQNtyCroGr7ou4QfYRenDbImtrocXL0mpc9xaEI1ISjdTtAeRtcqAlHi7b+Txew+RkNn0finyizyXDN8qpFrZ9FiIKwkxwdZYXog/VxuGY5Gr6U9SyUCXcWR6OkzS6bvx3RX0LwAtuNaNskAAAAAElFTkSuQmCC';
                try{
                    photoBase64 = await client.getPhotoAsync(tokenResponse.token);    
                }catch (error) {
                    console.error(`Error fetching photo: ${error}`);
                }
                
                const card = CardFactory.thumbnailCard("", CardFactory.images([photoBase64]));
                await stepContext.context.sendActivity({ attachments: [card] });
                return await stepContext.prompt(CONFIRM_PROMPT, 'Would you like to view your token?');
            } catch (error) {
                console.error(`Error: ${error}`);
                await stepContext.context.sendActivity(`An error occurred while processing your request. ${error}`);
                return await stepContext.endDialog();
            }
            
        }
    }

    /**
     * Ensures the OAuth token is available.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async ensureOAuth(stepContext) {
        await stepContext.context.sendActivity('Thank you.');

        const result = stepContext.result;
        if (result) {
            return await stepContext.beginDialog(OAUTH_PROMPT);
        }
        return await stepContext.endDialog();
    }

    /**
     * Displays the OAuth token to the user.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async displayToken(stepContext) {
        const tokenResponse = stepContext.result;
        if (tokenResponse && tokenResponse.token) {
            await stepContext.context.sendActivity(`Here is your token: ${tokenResponse.token}`);
        }
        return await stepContext.endDialog();
    }
}

module.exports.MainDialog = MainDialog;
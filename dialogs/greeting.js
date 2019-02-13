const { ComponentDialog, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');

const initialId = 'mainDialog';

class GreetingDialog extends ComponentDialog {
    constructor(id) {
        super(id);

        this.initialDialogId = initialId;

        this.addDialog(new TextPrompt('userNamePrompt'));

        this.addDialog(new WaterfallDialog(initialId, [
            async (step) => {
                const options = step.options;
                if (options.userProfile.name === undefined) {
                    return await step.prompt('userNamePrompt', 'Hello, What is your name.');
                } else {
                    step.context.sendActivity(`Hello ${options.userProfile.name}. How may i help you?`);
                    return await step.endDialog();
                }
            },
            async (step) => {
                step.context.sendActivity(`Hello ${step.result}. How may i help you?`);
                return await step.endDialog();
            },
        ]));
    }
}

exports.GreetingDialog = GreetingDialog;

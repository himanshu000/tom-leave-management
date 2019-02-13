const { ActionTypes, ActivityTypes, CardFactory } = require('botbuilder');

const {
    ComponentDialog,
    WaterfallDialog,
} = require('botbuilder-dialogs');

const initialId = 'mainDialog';

class FlexibleHolidaysDialog extends ComponentDialog {
    constructor(id) {
        super(id);

        this.initialDialogId = initialId;

        this.addDialog(new WaterfallDialog(initialId, [
            async (step) => {
                const flexibleHolidays = step.options.flexibleHolidays;
                if (flexibleHolidays.length > 0) {
                    const flexibleHolidaysButtons = [];
                    flexibleHolidays.forEach((flexibleHoliday) => {
                        flexibleHolidaysButtons.push({
                            title: `${flexibleHoliday.date.toDateString()} - ${flexibleHoliday.holiday}`,
                            type: ActionTypes.PostBack,
                            value: `I want to apply for flexible leave on ${flexibleHoliday.date.toDateString()}`,
                        });
                    });

                    const flexibleHolidaysCard = CardFactory.heroCard(
                        'Flexible Holidays',
                        undefined,
                        flexibleHolidaysButtons,
                    );

                    const reply = {
                        attachments: [flexibleHolidaysCard],
                        type: ActivityTypes.Message,
                    };
                    await step.context.sendActivity(reply);
                } else {
                    await step.context.sendActivity('No upcoming public holiday in specified search.');
                    await step.context.sendActivity('Please modify your search.');
                }

                return await step.endDialog();
            },
        ]));
    }
}

exports.FlexibleHolidaysDialog = FlexibleHolidaysDialog;
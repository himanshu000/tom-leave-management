const { CardFactory } = require('botbuilder');

const {
    ComponentDialog,
    WaterfallDialog,
} = require('botbuilder-dialogs');

const initialId = 'mainDialog';

class PublicHolidaysDialog extends ComponentDialog {
    constructor(id) {
        super(id);

        this.initialDialogId = initialId;

        this.addDialog(new WaterfallDialog(initialId, [
            async (step) => {
                const publicHolidays = step.options.publicHolidays;
                if (publicHolidays.length > 0) {
                    const publicHolidayCard = {
                        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                        body: [
                            {
                                columns: [
                                    {
                                        items: [
                                            {
                                                color: 'Accent',
                                                horizontalAlignment: 'Right',
                                                isSubtle: true,
                                                size: 'ExtraLarge',
                                                text: 'Upcoming Holidays',
                                                type: 'TextBlock',
                                            },
                                        ],
                                        type: 'Column',
                                        width: 'stretch',
                                    },
                                ],
                                separator: true,
                                spacing: 'Medium',
                                type: 'ColumnSet',
                            },
                            {
                                columns: [
                                    {
                                        items: [
                                            {
                                                color: 'Accent',
                                                size: 'ExtraLarge',
                                                spacing: 'None',
                                                text: 'Date',
                                                type: 'TextBlock',
                                            },
                                        ],
                                        type: 'Column',
                                        width: 1,
                                    },
                                    {
                                        items: [
                                            {
                                                color: 'Accent',
                                                horizontalAlignment: 'Right',
                                                size: 'ExtraLarge',
                                                spacing: 'None',
                                                text: 'Holiday',
                                                type: 'TextBlock',
                                            },
                                        ],
                                        type: 'Column',
                                        width: 1,
                                    },
                                ],
                                separator: true,
                                spacing: 'Medium',
                                type: 'ColumnSet',
                            },
                        ],
                        type: 'AdaptiveCard',
                        version: '1.0',
                    };

                    publicHolidays.forEach((publicHoliday) => {
                        publicHolidayCard.body.push({
                            columns: [
                                {
                                    items: [
                                        {
                                            spacing: 'Small',
                                            text: `${publicHoliday.date.toDateString()}`,
                                            type: 'TextBlock',
                                        },
                                    ],
                                    type: 'Column',
                                    width: 1,
                                },
                                {
                                    items: [
                                        {
                                            horizontalAlignment: 'Right',
                                            spacing: 'Small',
                                            text: `${publicHoliday.holiday}`,
                                            type: 'TextBlock',
                                            weight: 'Bolder',
                                        },
                                    ],
                                    type: 'Column',
                                    width: 1,
                                },
                            ],
                            spacing: 'Medium',
                            type: 'ColumnSet',
                        });
                    });
                    await step.context.sendActivity({
                        attachments: [CardFactory.adaptiveCard(publicHolidayCard)],
                    });
                } else {
                    await step.context.sendActivity('No upcoming public holiday in specified search.');
                    await step.context.sendActivity('Please modify your search.');
                }

                return await step.endDialog();
            },
        ]));
    }
}

exports.PublicHolidaysDialog = PublicHolidaysDialog;
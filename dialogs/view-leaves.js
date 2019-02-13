const { CardFactory } = require('botbuilder');
const { ComponentDialog, WaterfallDialog } = require('botbuilder-dialogs');

const initialId = 'mainDialog';

class ViewLeavesDialog extends ComponentDialog {
    constructor(id) {
        super(id);

        this.initialDialogId = initialId;

        this.addDialog(new WaterfallDialog(initialId, [
            async (step) => {
                const options = step.options;
                const startDate = options.startDate;
                const endDate = options.endDate;
                let requestedLeaves = options.requestedLeaves.slice();
                let requestedFlexibleLeaves = options.requestedFlexibleLeaves.slice();
                const inputText = options.text;
                const viewLeavesCard = {
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
                                            text: `Your Leaves`,
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
                                {
                                    items: [
                                        {
                                            color: 'Accent',
                                            horizontalAlignment: 'Right',
                                            size: 'ExtraLarge',
                                            spacing: 'None',
                                            text: 'Type',
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

                requestedLeaves = requestedLeaves.filter((element) => {
                    return new Date(element) >= new Date(startDate) && new Date(element) <= new Date(endDate);
                });

                requestedFlexibleLeaves = requestedFlexibleLeaves.filter((element) => {
                    return new Date(element.date) >= new Date(startDate) && new Date(element.date) <= new Date(endDate);
                });

                if (inputText.toLowerCase().includes('flexible') && requestedFlexibleLeaves.length > 0) {
                    this.setFlexibleLeaveAdaptiveCard(requestedFlexibleLeaves, viewLeavesCard);
                    await step.context.sendActivity({
                        attachments: [CardFactory.adaptiveCard(viewLeavesCard)],
                    });
                } else if (requestedLeaves.length > 0 || requestedFlexibleLeaves.length > 0) {
                    this.setFlexibleLeaveAdaptiveCard(requestedFlexibleLeaves, viewLeavesCard);
                    this.setLeaveAdaptiveCard(requestedLeaves, viewLeavesCard);
                    await step.context.sendActivity({
                        attachments: [CardFactory.adaptiveCard(viewLeavesCard)],
                    });
                } else {
                    await step.context.sendActivity('No data to display');
                }

                return await step.endDialog();
            },
        ]));
    }

    setFlexibleLeaveAdaptiveCard(requestedFlexibleLeaves, viewLeavesCard) {
        requestedFlexibleLeaves.forEach((element) => {
            viewLeavesCard.body.push({
                columns: [
                    {
                        items: [
                            {
                                spacing: 'Small',
                                text: `${new Date(element.date).toDateString()}`,
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
                                text: `${element.holiday}`,
                                type: 'TextBlock',
                                weight: 'Bolder',
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
                                text: `Flexible`,
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
    }

    setLeaveAdaptiveCard(requestedLeaves, viewLeavesCard) {
        requestedLeaves.forEach((element) => {
            viewLeavesCard.body.push({
                columns: [
                    {
                        items: [
                            {
                                spacing: 'Small',
                                text: `${new Date(element).toDateString()}`,
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
                                text: ``,
                                type: 'TextBlock',
                                weight: 'Bolder',
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
                                text: `Regular`,
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
    }
}

exports.ViewLeavesDialog = ViewLeavesDialog;
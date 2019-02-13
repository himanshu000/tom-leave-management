const { ActivityTypes } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');

const { GreetingDialog } = require('./dialogs/greeting');

const { FlexibleHolidaysDialog } = require('./dialogs/flexible-holidays');
const { PublicHolidaysDialog } = require('./dialogs/public-holidays');
const { SubmitLeaveDialog } = require('./dialogs/submit-leave');
const { ViewLeavesDialog } = require('./dialogs/view-leaves');

const GREETING_DIALOG = 'greetingDialog';
const PUBLIC_HOLIDAYS_DIALOG = 'publicHolidaysDialog';
const FLEXIBLE_HOLIDAYS_DIALOG = 'flexibleHolidaysDialog';
const SUBMIT_LEAVE_DIALOG = 'submitLeaveDialog';
const VIEW_LEAVE_DIALOG = 'viewLeaveDialog';

const DIALOG_STATE_PROPERTY = 'dialogState';
const USER_PROFILE_PROPERTY = 'userProfileProperty';

const GREETING_INTENT = 'greetings';
const PUBLIC_HOLIDAYS_INTENT = 'publicHolidays';
const FLEXIBLE_HOLIDAYS_INTENT = 'flexibleHolidays';
const SUBMIT_FLEXIBLE_LEAVE_INTENT = 'submitFlexibleLeave';
const SUBMIT_LEAVE_INTENT = 'submitLeave';
const VIEW_LEAVE_INTENT = 'viewLeaves';
const CANCEL_INTENT = 'cancel';
const HELP_INTENT = 'help';

const USER_NAME_ENTITIES = ['personName', 'personName_patternAny'];

class MyBot {
    constructor(application, luisPredictionOptions, conversationState, userState) {
        if (!conversationState) {
            throw new Error('Missing parameter.  conversationState is required');
        }
        if (!userState) {
            throw new Error('Missing parameter.  userState is required');
        }

        this.luisRecognizer = new LuisRecognizer(
            application,
            luisPredictionOptions,
            true,
        );

        this.userProfileAccessor = userState.createProperty(USER_PROFILE_PROPERTY);
        this.dialogState = conversationState.createProperty(DIALOG_STATE_PROPERTY);

        // Create top-level dialog(s)
        this.dialogs = new DialogSet(this.dialogState);
        // Add the Greeting dialog to the set
        this.dialogs.add(new GreetingDialog(GREETING_DIALOG));
        this.dialogs.add(new PublicHolidaysDialog(PUBLIC_HOLIDAYS_DIALOG));
        this.dialogs.add(new FlexibleHolidaysDialog(FLEXIBLE_HOLIDAYS_DIALOG));
        this.dialogs.add(new SubmitLeaveDialog(SUBMIT_LEAVE_DIALOG));
        this.dialogs.add(new ViewLeavesDialog(VIEW_LEAVE_DIALOG));

        this.conversationState = conversationState;
        this.userState = userState;
    }

    async onTurn(context) {
        if (context.activity.type === ActivityTypes.Message) {
            const results = await this.luisRecognizer.recognize(context);
            const topIntent = LuisRecognizer.topIntent(results);

            await this.updateUserProfile(results, context);

            let dialogResult;
            const dc = await this.dialogs.createContext(context);

            const interrupted = await this.isTurnInterrupted(dc, results);
            if (interrupted) {
                if (dc.activeDialog !== undefined) {
                    dialogResult = await dc.repromptDialog();
                }
            } else {
                dialogResult = await dc.continueDialog();
            }
            if (!dc.context.responded) {
                switch (dialogResult.status) {
                    case DialogTurnStatus.empty:
                        const userProfile = await this.userProfileAccessor.get(context);
                        switch (topIntent) {
                            case GREETING_INTENT: {
                                await dc.beginDialog(GREETING_DIALOG, { userProfile });
                                break;
                            }
                            case PUBLIC_HOLIDAYS_INTENT: {
                                const { startDate, endDate } = this.getDaysRange(results);

                                await dc.beginDialog(PUBLIC_HOLIDAYS_DIALOG, {
                                    publicHolidays: publicHolidays.filter((publicHoliday) => {
                                        return (publicHoliday.date >= startDate && publicHoliday.date <= endDate);
                                    }),
                                });
                                break;
                            }
                            case FLEXIBLE_HOLIDAYS_INTENT: {
                                const { startDate, endDate } = this.getDaysRange(results);

                                await dc.beginDialog(FLEXIBLE_HOLIDAYS_DIALOG, {
                                    flexibleHolidays: flexibleHolidays.filter((flexibleHoliday) => {
                                        return (flexibleHoliday.date >= startDate && flexibleHoliday.date <= endDate);
                                    }),
                                });
                                break;
                            }
                            case SUBMIT_FLEXIBLE_LEAVE_INTENT: {
                                const resolution = results.luisResult.entities[0].resolution;
                                const index = resolution.values.length - 1;
                                const flexibleLeaveDate = new Date(resolution.values[index].value);
                                if (userProfile.requestedFlexibleLeaves.length < 3) {
                                    if (userProfile.requestedFlexibleLeaves.findIndex((element) => {
                                        return new Date(element).toDateString() === flexibleLeaveDate.toDateString();
                                    }) > -1) {
                                        await dc.context.sendActivity(`You have already requested for flexible holiday on this day.`);
                                    } else {
                                        const flexibleLeave = flexibleHolidays.find((element) => {
                                            return new Date(element.date).toDateString() === flexibleLeaveDate.toDateString();
                                        });
                                        userProfile.requestedFlexibleLeaves.push(flexibleLeave);
                                        await dc.context.sendActivity(`Your flexible leave request has been submitted.`);
                                        await this.userProfileAccessor.set(context, userProfile);
                                        await this.userState.saveChanges(context);
                                    }
                                } else {
                                    await dc.context.sendActivity(`You have already availed maximum flexible leave`);
                                }
                                break;
                            }
                            case VIEW_LEAVE_INTENT: {
                                const { startDate, endDate } = this.getDaysRange(results);

                                await dc.beginDialog(VIEW_LEAVE_DIALOG, {
                                    endDate,
                                    requestedFlexibleLeaves: userProfile.requestedFlexibleLeaves,
                                    requestedLeaves: userProfile.requestedLeaves,
                                    startDate,
                                    text: results.text,
                                });
                                break;
                            }
                            case SUBMIT_LEAVE_INTENT: {
                                let startDate;
                                let endDate;
                                if (results.luisResult && results.luisResult.entities && results.luisResult.entities.length > 0) {
                                    const resolution = results.luisResult.entities[0].resolution;
                                    const index = resolution.values.length - 1;
                                    startDate = resolution.values[index].start ? new Date(resolution.values[index].start) :
                                        new Date(resolution.values[index].value);
                                    endDate = resolution.values[index].end ? new Date(resolution.values[index].end) : undefined;
                                }
                                await dc.beginDialog(SUBMIT_LEAVE_DIALOG, {
                                    endDate,
                                    publicHolidays,
                                    startDate,
                                    userProfileAccessor: this.userProfileAccessor,
                                });
                                await this.userState.saveChanges(context);
                                break;
                            }
                            default: {
                                await dc.context.sendActivity(`I didn't understand what you just said to me.`);
                                break;
                            }
                        }
                        break;
                    case DialogTurnStatus.waiting:
                        // The active dialog is waiting for a response from the user, so do nothing.
                        break;
                    case DialogTurnStatus.complete:
                        // All child dialogs have ended. so do nothing.
                        break;
                    default:
                        // Unrecognized status from child dialog. Cancel all dialogs.
                        await dc.cancelAllDialogs();
                        break;
                    }
            }
        } else if (context.activity.type === ActivityTypes.ConversationUpdate) {
            await this.sendWelcomeMessage(context);
        }

        // make sure to persist state at the end of a turn.
        await this.conversationState.saveChanges(context);
        await this.userState.saveChanges(context);
    }

    getDaysRange(results) {
        let startDate;
        let endDate;
        if (results.luisResult && results.luisResult.entities && results.luisResult.entities.length > 0) {
            const resolution = results.luisResult.entities[0].resolution;
            const index = resolution.values.length - 1;
            startDate = new Date(resolution.values[index].start);
            endDate = new Date(resolution.values[index].end);
        } else {
            const date = new Date();
            startDate = new Date(date.getFullYear(), date.getMonth(), 1);
            endDate = new Date('12/31/2019');
        }
        return { startDate, endDate };
    }

    // Sends welcome messages to conversation members when they join the conversation.
    // Messages are only sent to conversation members who aren't the bot.
    async sendWelcomeMessage(turnContext) {
        // If any new membmers added to the conversation
        if (turnContext.activity.membersAdded) {
            const replyPromises = turnContext.activity.membersAdded.map(async (member) => {
                if (member.id !== turnContext.activity.recipient.id) {
                    const message = `Welcome to Nagarro Leave Management.`;
                    await turnContext.sendActivity(message);
                }
            });
            await Promise.all(replyPromises);
        }
    }

    /**
     * Look at the LUIS results and determine if we need to handle
     * an interruptions due to a Help or Cancel intent
     *
     * @param {DialogContext} dc - dialog context
     * @param {LuisResults} luisResults - LUIS recognizer results
     */
    async isTurnInterrupted(dc, luisResults) {
        const topIntent = LuisRecognizer.topIntent(luisResults);

        // see if there are any conversation interrupts we need to handle
        if (topIntent === CANCEL_INTENT) {
            if (dc.activeDialog) {
                // cancel all active dialog (clean the stack)
                await dc.cancelAllDialogs();
                await dc.context.sendActivity(`Ok.  I've cancelled our last activity.`);
            } else {
                await dc.context.sendActivity(`I don't have anything to cancel.`);
            }
            return true; // this is an interruption
        }

        if (topIntent === HELP_INTENT) {
            await dc.context.sendActivity(`Let me try to provide some help.`);
            await dc.context.sendActivity(`I understand greetings, being asked for help, or being asked to cancel what I am doing.`);
            return true; // this is an interruption
        }
        return false; // this is not an interruption
    }

    /**
     * Helper function to update user profile with entities returned by LUIS.
     *
     * @param {LuisResults} luisResults - LUIS recognizer results
     * @param {DialogContext} dc - dialog context
     */
    async updateUserProfile(luisResult, context) {
        // get userProfile object using the accessor
        let userProfile = await this.userProfileAccessor.get(context);
        if (userProfile === undefined) {
            userProfile = {};
            userProfile.requestedFlexibleLeaves = [];
            userProfile.requestedLeaves = [];
        }

        // Do we have any entities?
        if (Object.keys(luisResult.entities).length !== 1) {
            // see if we have any user name entities
            USER_NAME_ENTITIES.forEach((element) => {
                if (luisResult.entities[element] !== undefined) {
                    const lowerCaseName = luisResult.entities[element][0];

                    // capitalize and set user name
                    userProfile.name = lowerCaseName.charAt(0).toUpperCase() + lowerCaseName.substr(1);
                }
             });
        }

        await this.userProfileAccessor.set(context, userProfile);
        await this.userState.saveChanges(context);
    }
}

exports.MyBot = MyBot;

const publicHolidays = [
    {date: new Date('01/01/2019'), holiday: 'New Year\'s Day'},
    {date: new Date('01/26/2019'), holiday: 'Republic Day'},
    {date: new Date('02/10/2019'), holiday: 'Vasant Panchmi'},
    {date: new Date('03/21/2019'), holiday: 'Holi'},
    {date: new Date('04/13/2019'), holiday: 'Ram Navmi'},
    {date: new Date('04/14/2019'), holiday: 'Vaisakhi'},
    {date: new Date('08/15/2019'), holiday: 'Independence Day and Raksha Bandhan'},
    {date: new Date('04/24/2019'), holiday: 'Janmashtami'},
    {date: new Date('10/02/2019'), holiday: 'Gandhi Jayanti'},
    {date: new Date('10/08/2019'), holiday: 'Dussehra'},
    {date: new Date('10/27/2019'), holiday: 'Diwali'},
    {date: new Date('10/28/2019'), holiday: 'Diwali'},
    {date: new Date('12/25/2019'), holiday: 'Christmas'},
];

const flexibleHolidays = [
    {date: new Date('01/14/2019'), holiday: 'Makar Sankranti'},
    {date: new Date('01/15/2019'), holiday: 'Pongal'},
    {date: new Date('03/04/2019'), holiday: 'Maha Shivratri'},
    {date: new Date('04/19/2019'), holiday: 'Good Friday'},
    {date: new Date('05/24/2019'), holiday: 'Nagarro\'s Day of Reason' },
    {date: new Date('06/05/2019'), holiday: 'Idul Fitr'},
    {date: new Date('08/12/2019'), holiday: 'Idul Juha'},
    {date: new Date('09/02/2019'), holiday: 'Ganesh Chaturthi'},
    {date: new Date('09/11/2019'), holiday: 'Onam'},
    {date: new Date('10/29/2019'), holiday: 'Bhai Dooj'},
    {date: new Date('11/12/2019'), holiday: 'Guru Nanak Jayanti'},
];

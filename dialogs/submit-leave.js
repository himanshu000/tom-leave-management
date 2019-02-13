const { ComponentDialog, WaterfallDialog } = require('botbuilder-dialogs');

const initialId = 'mainDialog';

class SubmitLeaveDialog extends ComponentDialog {
    constructor(id) {
        super(id);

        this.initialDialogId = initialId;

        this.addDialog(new WaterfallDialog(initialId, [
            async (step) => {
                const options = step.options;
                const startDate = options.startDate;
                const endDate = options.endDate;
                const publicHolidays = options.publicHolidays;
                const userProfile = await options.userProfileAccessor.get(step.context);
                const requestedLeaves = userProfile.requestedLeaves;
                if (startDate && endDate) {
                    const end = new Date(endDate);
                    const datesToApplyLeave = [];
                    for (const d = new Date(startDate); d <= end; d.setDate(d.getDate() + 1)) {
                        const dayOfWeek = d.getDay();

                        if (dayOfWeek !== 6 && dayOfWeek !== 0 && publicHolidays.findIndex((element) => {
                            return new Date(element.date).toDateString() !== d.toDateString();
                        }) > -1 && userProfile.requestedLeaves.findIndex((element) => {
                            return new Date(element).toDateString() !== d.toDateString();
                        }) > -1) {
                            datesToApplyLeave.push(d);
                        }

                        if (requestedLeaves.length + datesToApplyLeave.length <= 27) {
                            userProfile.requestedLeaves.push(datesToApplyLeave);
                            await options.userProfileAccessor.set(step.context, userProfile);
                            await step.context.sendActivity(`Your leave has been applied`);
                        } else {
                            step.context.sendActivity('You have already exhausted your leaves');
                        }
                    }
                } else if (startDate) {
                    const dateOfLeave = new Date(startDate);
                    await this.availSingleDayLeave(dateOfLeave, step, publicHolidays, userProfile, requestedLeaves);
                } else {
                    const currDate = new Date();
                    await this.availSingleDayLeave(currDate, step, publicHolidays, userProfile, requestedLeaves);
                }
                return await step.endDialog();
            },
        ]));
    }

    async availSingleDayLeave(currDate, step, publicHolidays, userProfile, requestedLeaves) {
        const dayOfWeek = currDate.getDay();
        if (dayOfWeek === 6 || dayOfWeek === 0) {
            step.context.sendActivity(`Your leave has been applied`);
        } else if (publicHolidays.findIndex((element) => {
            return new Date(element.date).toDateString() === currDate.toDateString();
        }) > -1) {
            step.context.sendActivity(`Your leave has been applied`);
        } else if (userProfile.requestedLeaves.findIndex((element) => {
            return new Date(element).toDateString() === currDate.toDateString();
        }) > -1) {
            step.context.sendActivity(`Your leave has been applied`);
        } else {
            if (requestedLeaves.length < 27) {
                userProfile.requestedLeaves.push(currDate);
                await step.options.userProfileAccessor.set(step.context, userProfile);
                await step.context.sendActivity(`Your leave has been applied`);
            } else {
                step.context.sendActivity('You have already exhausted your leaves');
            }
        }
    }
}

exports.SubmitLeaveDialog = SubmitLeaveDialog;
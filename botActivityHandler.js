// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes,
    TeamsInfo
} = require('botbuilder');

const DEBOUNCE_THRESHOLD_MILLIS = 1000 * 60;

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
        /* Conversation Bot */
        /*  Teams bots are Microsoft Bot Framework bots.
            If a bot receives a message activity, the turn handler sees that incoming activity
            and sends it to the onMessage activity handler.
            Learn more: https://aka.ms/teams-bot-basics.

            NOTE:   Ensure the bot endpoint that services incoming conversational bot queries is
                    registered with Bot Framework.
                    Learn more: https://aka.ms/teams-register-bot. 
        */
        // Registers an activity event handler for the message event, emitted for every incoming message activity.
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            switch (context.activity.text.trim().toLowerCase()) {
            case 'choose an order':
                await this.orderActivityAsync(context);
                break;
            default:
                // By default for unknown activity sent by user show
                // a card with the available actions.
                const value = { count: 0 };
                const card = CardFactory.heroCard(
                    'I can choose an order for your standup meeting.',
                    null,
                    [{
                        type: ActionTypes.MessageBack,
                        title: 'Choose an order',
                        value: value,
                        text: 'Choose an order'
                    }]);
                await context.sendActivity({ attachments: [card] });
                break;
            }
            await next();
        });
        /* Conversation Bot */

        this.recentOrderings = {};
    }

    /* Conversation Bot */
    /**
     * Say hello and @ mention the current user.
     */
    async orderActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const recentRequester = this.maybeGetRecentRequester(context);
        var replyActivity;
        if (recentRequester) {
            replyActivity = MessageFactory.text(`Sorry ${mention.text }, ${recentRequester.name} beat you to it.`);
        } else {
            const members = await TeamsInfo.getMembers(context);
            const memberNamesInOrder = this.orderMemberNames(members);

            replyActivity = MessageFactory.text(`Hi ${ mention.text }, here is a random order:\n\n${memberNamesInOrder.join("\n\n")}`);
        }
        
        replyActivity.entities = [mention];
        
        await context.sendActivity(replyActivity);
    }
    /* Conversation Bot */

    orderMemberNames(members) {
        const givenNameCount = {};
        members.forEach(member => {
            if (givenNameCount.hasOwnProperty(member.givenName)) {
                givenNameCount[member]++;
            } else {
                givenNameCount[member] = 1;
            }
        });
        
        const displayNames = members.map(member => {
            if (!member.givenName) {
                return member.name;
            }

            return `${member.givenName} ${givenNameCount[member.givenName] > 1 ? member.surname : ''}`;
        })
        .map(name => name.trim());
;

        return displayNames
            .map(name => name.startsWith("Brian") ? `🎉🎈✨${name}✨🎈🎉` : name)
            .map(name => ({sort: Math.random(), value: name}))
            .sort((a, b) => a.sort - b.sort)
            .map(a => a.value);
    }

    maybeGetRecentRequester(context) {
        const convKey = context.activity.conversation.id;
        if (this.recentOrderings[convKey] && (Date.now() - this.recentOrderings[convKey].timestamp < DEBOUNCE_THRESHOLD_MILLIS)) {
            return this.recentOrderings[convKey].requester;
        }

        this.recentOrderings[convKey] = {
            timestamp: Date.now(),
            requester: context.activity.from
        };
        return undefined;
    }

}

module.exports.BotActivityHandler = BotActivityHandler;


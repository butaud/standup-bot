// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
  TurnContext,
  MessageFactory,
  TeamsActivityHandler,
  CardFactory,
  ActionTypes,
  TeamsInfo,
  TeamsChannelAccount,
} from "botbuilder";

const DEBOUNCE_THRESHOLD_MILLIS = 1000 * 60;

type RecentOrdering = {
  requester: string;
  timestamp: number;
};

class BotActivityHandler extends TeamsActivityHandler {
  private readonly recentOrderings: Record<string, RecentOrdering>;
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
        case "choose an order":
          await this.orderActivityAsync(context);
          break;
        default:
          // By default for unknown activity sent by user show
          // a card with the available actions.
          const value = { count: 0 };
          const card = CardFactory.heroCard(
            "I can choose an order for your standup meeting.",
            undefined,
            [
              {
                type: ActionTypes.MessageBack,
                title: "Choose an order",
                value: value,
                text: "Choose an order",
              },
            ]
          );
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
  async orderActivityAsync(context: TurnContext) {
    const TextEncoder = require("html-entities").XmlEntities;

    const mention = {
      mentioned: context.activity.from,
      text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
      type: "mention",
    };

    const recentRequester = this.maybeGetRecentRequester(context);
    var replyActivity;
    if (recentRequester) {
      replyActivity = MessageFactory.text(
        `Sorry ${mention.text}, ${recentRequester} beat you to it.`
      );
    } else {
      await context.sendActivity(
        MessageFactory.text("Sure, it will take just a minute.")
      );
      const members = (await TeamsInfo.getPagedMembers(context)).members;
      const membersInOrder = this.orderMembers(members);

      let memberNamesInOrder: string[] = [];

      if (
        context.activity.conversation.tenantId &&
        context.activity.channelData?.meeting?.id
      ) {
        const membersMeetingPresence = await this.getMeetingPresence(
          context,
          context.activity.conversation.tenantId,
          context.activity.channelData.meeting.id,
          membersInOrder.map((member) => member.id)
        );
        const presentMembers = membersInOrder.filter(
          (member) => membersMeetingPresence[member.id]
        );
        const absentMembers = membersInOrder.filter(
          (member) => !membersMeetingPresence[member.id]
        );

        memberNamesInOrder = [
          ...this.formatMemberNames(presentMembers).map(
            (name) => `**${name}**`
          ),
          ...this.formatMemberNames(absentMembers),
        ];
      } else {
        memberNamesInOrder = this.formatMemberNames(membersInOrder);
      }

      replyActivity = MessageFactory.text(
        `${
          mention.text
        }, here is the random order you requested:\n\n${memberNamesInOrder.join(
          "\n\n"
        )}`
      );
    }

    replyActivity.entities = [mention];

    await context.sendActivity(replyActivity);
  }
  /* Conversation Bot */

  async getMeetingPresence(
    context: TurnContext,
    tenantId: string,
    meetingId: string,
    memberIds: string[]
  ): Promise<Record<string, boolean>> {
    return await memberIds.reduce(async (prev, memberId) => {
      const prevDict = await prev;
      try {
        prevDict[memberId] =
          (
            await TeamsInfo.getMeetingParticipant(
              context,
              meetingId,
              memberId,
              tenantId
            )
          ).meeting?.inMeeting ?? false;
      } catch (e) {
        console.error(e);
      }
      return prevDict;
    }, Promise.resolve({} as Record<string, boolean>));
  }

  formatMemberNames(members: TeamsChannelAccount[]) {
    const givenNameCount: Record<string, number> = {};
    members.forEach((member) => {
      if (member.givenName) {
        if (givenNameCount.hasOwnProperty(member.givenName)) {
          givenNameCount[member.givenName]++;
        } else {
          givenNameCount[member.givenName] = 1;
        }
      }
    });

    const displayNames = members
      .map((member) => {
        if (!member.givenName) {
          return member.name;
        }

        return `${member.givenName} ${
          givenNameCount[member.givenName] > 1 ? member.surname : ""
        }`;
      })
      .map((name) => name.trim());
    return displayNames;
  }

  orderMembers(members: TeamsChannelAccount[]) {
    return members
      .map((member) => ({ sort: Math.random(), value: member }))
      .sort((a, b) => a.sort - b.sort)
      .map((a) => a.value);
  }

  maybeGetRecentRequester(context: TurnContext) {
    const convKey = context.activity.conversation.id;
    if (
      this.recentOrderings[convKey] &&
      Date.now() - this.recentOrderings[convKey].timestamp <
        DEBOUNCE_THRESHOLD_MILLIS
    ) {
      return this.recentOrderings[convKey].requester;
    }

    this.recentOrderings[convKey] = {
      timestamp: Date.now(),
      requester: context.activity.from.name,
    };
    return undefined;
  }
}

module.exports.BotActivityHandler = BotActivityHandler;

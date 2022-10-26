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
        case "help":
          const helpMessage1 = MessageFactory.text(
            "I can choose a random order for all the people in your meeting, making sure that anyone actively on the call is sorted before anyone not on the call. You don't need to use any special command: I'll choose an order if you say anything but \"help\"."
          );
          await context.sendActivity(helpMessage1);
          const helpMessage2 = MessageFactory.text(
            'If multiple people are joining from a conference room or someone is dialing in from a phone, you can tell me about them using "override:alias". You can add as many of these as you want, separated by spaces. For example: "override:bob override:mary"'
          );
          await context.sendActivity(helpMessage2);
          break;
        default:
          await this.orderActivityAsync(context);
          break;
      }
      await next();
    });
    /* Conversation Bot */

    this.recentOrderings = {};
  }

  getOverrideEmails(inputText: string): Set<string> {
    const tokens = inputText.split(" ");
    return new Set(
      tokens
        .filter((token) => token.startsWith("override:"))
        .map((token) => token.split(":")[1])
        // filter out any malformed (nothing past the ":")
        .filter((alias) => alias)
        .map((alias) => `${alias.toLowerCase()}@microsoft.com`)
    );
  }

  /* Conversation Bot */
  /**
   * Say hello and @ mention the current user.
   */
  async orderActivityAsync(context: TurnContext) {
    const inputText = context.activity.text.trim();
    const overrideEmails = this.getOverrideEmails(inputText);

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
        MessageFactory.text(
          `${context.activity.from.name}, I'm choosing an order for your meeting. It will take just a minute.`
        )
      );
      const members = (await TeamsInfo.getPagedMembers(context)).members;
      const membersInOrder = this.orderMembers(members);

      const membersMeetingPresence =
        context.activity.conversation.tenantId &&
        context.activity.channelData?.meeting?.id
          ? await this.getMeetingPresence(
              context,
              context.activity.conversation.tenantId,
              context.activity.channelData.meeting.id,
              membersInOrder.map((member) => member.id)
            )
          : {};

      const presentMembers: TeamsChannelAccount[] = [];
      const absentMembers: TeamsChannelAccount[] = [];

      overrideEmails.forEach((email) =>
        console.log(`override email: ${email}`)
      );
      membersInOrder.forEach((member) => {
        const memberEmail = member.email?.toLowerCase();
        console.log(`${memberEmail}`);
        if (
          membersMeetingPresence[member.id] ||
          (memberEmail && overrideEmails.has(memberEmail))
        ) {
          presentMembers.push(member);
        } else {
          absentMembers.push(member);
        }
      });

      const memberNamesInOrder = [
        ...this.formatMemberNames(presentMembers).map((name) => `**${name}**`),
        ...this.formatMemberNames(absentMembers),
      ];

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

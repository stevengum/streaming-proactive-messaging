import { ActivityHandler, MessageFactory } from 'botbuilder';
import { QueueClient } from '@azure/storage-queue';

export class EchoBot extends ActivityHandler {
    constructor(private readonly queueMap: Record<string, string>, private readonly queueConnectionString: string) {
        super();
        this.onEvent(async (context, next) => {
            // On the custom event from Web Chat, populate the queueMap with the queueId and conversationId to know which queue to upload messages to.
            if (context.activity.name === 'webchat/join') {
                const { queueId } = context.activity.value;
                this.queueMap[queueId] = context.activity.conversation.id;
                this.queueMap[context.activity.conversation.id] = queueId;
            }

            await next();
        });

        this.onMessage(async (context, next) => {
            const replyText = `Echo: ${ context.activity.text }`;
            await context.sendActivity(MessageFactory.text(replyText));

            // Get the queueId to create a QueueClient and create messages to upload.
            const queueId = this.queueMap[context.activity.conversation.id];
            const client = new QueueClient(this.queueConnectionString, queueId);
            const otherBotReply = `You said "${ context.activity.text }" when talking with SkillA bot.`;
            
            // Upload messages to the appropriate Azure Storage Queue.
            await client.sendMessage(JSON.stringify({ text: otherBotReply, speak: otherBotReply }));
            
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            await next();
        });
    }
}

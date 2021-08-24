import fetch from 'cross-fetch';
import { ActivityHandler, CloudAdapterBase, ConversationReference, TurnContext } from 'botbuilder';
import { ConnectorClient, ConnectorFactory } from 'botframework-connector';
import { QueueClient } from '@azure/storage-queue';

export class EchoBot extends ActivityHandler {
    constructor(private readonly queueMap: Record<string, [ConversationReference, ConnectorFactory, ConnectorClient]>, private readonly queueConnectionString: string, private readonly webChatBotEndpoint) {
        super();

        this.onMessage(async (context, next) => {
            await context.sendActivity(`You sent '${ context.activity.text }'`);
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;

            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    const welcomeMessage = 'Navigate to http://localhost:3978/api/dequeue to proactively message everyone who has previously messaged this bot.';
                    await context.sendActivity(welcomeMessage);

                    // Instantiate a QueueClient which will be used to create and manipulate a queue
                    const queueId = `directlinespeech-${context.activity.conversation.id}`;
                    const queueClient = new QueueClient(this.queueConnectionString, queueId);
        
                    // Create the queue and populate queueMap
                    await queueClient.create();
                    console.info(`Queue "${queueId}" created.`);
                    const conversationReference = TurnContext.getConversationReference(context.activity) as ConversationReference;
                    const connectorClient = context.turnState.get(context.adapter.ConnectorClientKey);
                    const connectorFactory = context.turnState.get((context.adapter as CloudAdapterBase).ConnectorFactoryKey);
                    this.queueMap[queueId] = [conversationReference, connectorFactory, connectorClient];
        
                    // POST queueId and relevant context to other bot using Web Chat
                    await fetch(`${this.webChatBotEndpoint}/newconversation/${conversationReference.conversation.id}`, {
                        method: 'POST',
                        body: JSON.stringify({ queueId, context: { conversationReference, firstMessage: context.activity.text } })
                    });

                    // Send message to user with link to open Web Chat with other bot
                    await context.sendActivity(`Go to ${this.webChatBotEndpoint}/webchat?queueId=${queueId} to use a synchronized Web Chat client with another bot.`);
                }
            }
            await next();
        });
    }
}

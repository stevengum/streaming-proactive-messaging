import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';

import { CloudAdapter, ConfigurationBotFrameworkAuthentication, ConversationReference } from 'botbuilder';
import { ConnectorClient, ConnectorFactory } from 'botframework-connector';
import { EchoBot } from './bot';
import { QueueClient } from '@azure/storage-queue';

const ENV_FILE = path.join(__dirname, '..', '.env');
config({ path: ENV_FILE });

const server = restify.createServer();
const port = process.env.port || process.env.PORT || 3978;
server.listen(port, () => {
    console.log(`\n${server.name} listening to http://localhost:3978`);
    console.log(`\nTo talk to this bot, navigate to http://localhost:3978`);
});

server.use(restify.plugins.bodyParser({
    mapParams: true
}));

const auth = new ConfigurationBotFrameworkAuthentication(process.env as any)
const adapter = new CloudAdapter(auth);

const onTurnErrorHandler = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

adapter.onTurnError = onTurnErrorHandler;

const queueIdMap: Record<string, [ConversationReference, ConnectorFactory, ConnectorClient]> = {};
const myBot = new EchoBot(queueIdMap, process.env.AZURE_STORAGE_CONNECTION_STRING, process.env.WEB_CHAT_BOT_ENDPOINT);

server.post('/api/messages', (req, res) => {
    adapter.process(req, res, async (context) => {
        await myBot.run(context);
    });
});
server.on('upgrade', (req, socket, head) => {
    adapter.process(req, socket as any, head, async (context) => await myBot.run(context));
});
server.opts('/api/messages', (req, res) => {
    res.send(200);
});

server.get('/', restify.plugins.serveStatic({
    directory: path.join(__dirname, '../public'),
    file: 'index.html'
}));

// Using an endpoint to trigger a fetch of messages from Azure Queue Storage.
// In production-level scenarios, a background task would run on set intervals
// to dequeue messages and proactively send messages to the end users over
// streaming connections.
server.get('/api/dequeue', async (_, res) => {
    const messagesToSend = Object.entries(queueIdMap).map(async ([queueId, [convRef, connectorFactory, connectorClient]]) => {
        return new Promise<void>(async (resolve, _reject) => {
            const queueClient = new QueueClient(process.env.AZURE_STORAGE_CONNECTION_STRING, queueId);
            const { receivedMessageItems: items } = await queueClient.receiveMessages();
            console.log(`Number of items dequeued: ${items.length}`);

            if (!items.length) {
                resolve();
            } else {
                resolve(adapter.continueConversationAsync(process.env.MicrosoftAppId, convRef, async (context) => {
                    context.turnState.set(adapter.ConnectorClientKey, connectorClient);
                    context.turnState.set(adapter.ConnectorFactoryKey, connectorFactory);
    
                    await Promise.all(
                        items.map(
                            (item) => new Promise(
                                (innerResolve, _innerReject) => {
                                    const proactiveMessage = JSON.parse(item.messageText);
                                    innerResolve(context.sendActivity(proactiveMessage));
                            })
                        )
                    );
                    })
                );
            }
        });
    });

    try {
        await Promise.all(messagesToSend);
        res.setHeader('Content-Type', 'text/html');
        res.writeHead(200);
        res.write('<html><body><h1>Proactive messages have been sent.</h1></body></html>');
        res.end();
    } catch (err) {
        console.error(err);

        res.writeHead(500);
        res.write(`<html><body><h1>Error dequeueing messages: "${(err as Error).message}"</h1></body></html>`);
        res.end();
    }
});

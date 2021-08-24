import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';

import { CloudAdapter, ConfigurationBotFrameworkAuthentication } from 'botbuilder';
import { EchoBot } from './bot';

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

const conversationReferences = {};
const myBot = new EchoBot(conversationReferences);

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

server.get('/api/notify', async (req, res) => {
    for (const conversationReference of Object.values(conversationReferences)) {
        await adapter.continueConversation(conversationReference, async turnContext => {
            await turnContext.sendActivity('proactive hello');
        });
    }
    res.setHeader('Content-Type', 'text/html');
    res.writeHead(200);
    res.write('<html><body><h1>Proactive messages have been sent.</h1></body></html>');
    res.end();
});

server.post('/api/notify', async (req, res) => {
    for (var prop in req.body) {
        var msg = req.body[prop];
        for (const conversationReference of Object.values(conversationReferences)) {
            await adapter.continueConversation(conversationReference, async turnContext => {
                await turnContext.sendActivity(msg);
            });
        }
    }
    res.setHeader('Content-Type', 'text/html');
    res.writeHead(200);
    res.write('Proactive messages have been sent.');
    res.end();
});
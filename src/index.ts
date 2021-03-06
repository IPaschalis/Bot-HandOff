import { MongooseProvider, mongoose } from "./mongoose-provider";
import { Handoff, ConversationState } from "./handoff";
import { commandsMiddleware } from "./commands";
import * as express from "express";
import * as bodyParser from "body-parser";
import * as teams from "botbuilder-teams";
import * as builder from "botbuilder";
import * as config from "config";
// import * as cors from 'cors';
const appInsights = require("applicationinsights");
let handoff;
// tslint:disable-next-line: variable-name
let support_address = null;

/**
 * @param connector: to get the team name from connector.fetchTeamInfo
 */
const setup = (bot, app, connector, isAgent, options) => {

    let mongooseProvider = null;
    let _retainData = null;
    let _directLineSecret = null;
    let _mongodbProvider = null;
    let _textAnalyticsKey = null;
    let _appInsightsInstrumentationKey = null;
    let _customerStartHandoffCommand = null;
    let _supportTeamId = null;
    let _supportChannelId = null;

    handoff = new Handoff(bot, connector, isAgent);

    options = options || {};

    if (!options.mongodbProvider && !process.env.MONGODB_PROVIDER) {
        throw new Error("Bot-Handoff: Mongo DB Connection String was not provided in setup options (mongodbProvider) or in the environment variables (MONGODB_PROVIDER)");
    } else {
        _mongodbProvider = options.mongodbProvider || process.env.MONGODB_PROVIDER;
        mongooseProvider = new MongooseProvider();
        mongoose.connect(_mongodbProvider, {useNewUrlParser: true});
    }

    if (!options.directlineSecret && !process.env.MICROSOFT_DIRECTLINE_SECRET) {
        throw new Error("Bot-Handoff: Microsoft Bot Builder Direct Line Secret was not provided in setup options (directlineSecret) or in the environment variables (MICROSOFT_DIRECTLINE_SECRET)");
    } else {
        _directLineSecret = options.directlineSecret || process.env.MICROSOFT_DIRECTLINE_SECRET;
    }

    if (!options.textAnalyticsKey && !process.env.CG_SENTIMENT_KEY) {
        // console.warn('Bot-Handoff: Microsoft Cognitive Services Text Analytics Key was not provided in setup options (textAnalyticsKey) or in the environment variables (CG_SENTIMENT_KEY). Sentiment will not be analysed in the transcript, the score will be recorded as -1 for all text.');
    } else {
        _textAnalyticsKey = options.textAnalyticsKey || process.env.CG_SENTIMENT_KEY;
        exports._textAnalyticsKey = _textAnalyticsKey;
    }

    if (!options.appInsightsInstrumentationKey && !process.env.APPINSIGHTS_INSTRUMENTATIONKEY) {
        // console.warn('Bot-Handoff: Microsoft Application Insights Instrumentation Key was not provided in setup options (appInsightsInstrumentationKey) or in the environment variables (APPINSIGHTS_INSTRUMENTATIONKEY). The conversation object will not be logged to Application Insights.');
    } else {
        _appInsightsInstrumentationKey = options.appInsightsInstrumentationKey || process.env.APPINSIGHTS_INSTRUMENTATIONKEY;
        appInsights.setup(_appInsightsInstrumentationKey).start();
        exports._appInsights = appInsights;
    }

    if (!options.retainData && !process.env.RETAIN_DATA) {
        console.warn('Bot-Handoff: Retain data value was not provided in setup options (retainData) or in the environment variables (RETAIN_DATA). Not providing this value or setting it to "false" means that if a customer speaks to an agent, the conversation record with that customer will be deleted after an agent disconnects the conversation. Set to "true" to keep all data records in the mongo database.');
    } else {
        _retainData = options.retainData || process.env.RETAIN_DATA;
        exports._retainData = _retainData;
    }

    if (!options.customerStartHandoffCommand && !process.env.CUSTOMER_START_HANDOFF_COMMAND) {
        console.warn("Bot-Handoff: The customer command to start the handoff was not provided in setup options (customerStartHandoffCommand) or in the environment variables (CUSTOMER_START_HANDOFF_COMMAND). The default command will be set to help. Regex is used on this command to make sure the activation of the handoff only works if the user types the exact phrase provided in this property.");
        _customerStartHandoffCommand = "help";
        exports._customerStartHandoffCommand = _customerStartHandoffCommand;
    } else {
        _customerStartHandoffCommand = options.customerStartHandoffCommand || process.env.CUSTOMER_START_HANDOFF_COMMAND;
        exports._customerStartHandoffCommand = _customerStartHandoffCommand;
    }

    if (!options.supportTeamId && !process.env.SUPPORT_TEAM_ID) {
        console.warn("Bot-Handoff: No support Team Id entered.");
    } else {
        _supportTeamId = options.supportTeamId || process.env.SUPPORT_TEAM_ID;
        exports._supportTeamId = _supportTeamId;
    }

    if (!options.supportChannelId && !process.env.SUPPORT_CHANNEL_ID) {
        console.warn("Bot-Handoff: No support Team Id entered.");
    } else {
        _supportChannelId = options.supportChannelId || process.env.SUPPORT_CHANNEL_ID;
        exports._supportChannelId = _supportChannelId;
    }

    support_address = {
        channelId: "msteams",
        bot: {
            id: config.MICROSOFT_APP_ID,
            name: "MareraBot",
        },
        conversation: {isGroup: true, id: _supportChannelId || _supportTeamId},
        serviceUrl: "https://smba.trafficmanager.net/emea-client-ss.msg/",
    };
    exports.support_address = support_address;

    if (bot) {
        bot.use(
            commandsMiddleware(bot, handoff),
            handoff.routingMiddleware(),
        );
    }

    if (app && _directLineSecret != null) {
        app.use(bodyParser.json());

        //// Create endpoint for agent / call center
        // app.use('/webchat', express.static('public'));

        // Endpoint to get current conversations
        app.get("/api/conversations", async (req, res) => {
            const authHeader = req.headers.authorization;
            console.log(authHeader);
            console.log(req.headers);
            if (authHeader) {
                if (authHeader === "Bearer " + _directLineSecret) {
                    const conversations = await mongooseProvider.getCurrentConversations();
                    res.status(200).send(conversations);
                }
            }
            res.status(401).send("Not Authorized");
        });

        // Endpoint to trigger handover
        app.post("/api/conversations", async (req, res) => {
            const authHeader = req.headers.authorization;
            console.log(authHeader);
            console.log(req.headers);
            if (authHeader) {
                if (authHeader === "Bearer " + _directLineSecret) {
                    if (await handoff.queueCustomerForAgent({ customerConversationId: req.body.conversationId })) {
                        res.status(200).send({ code: 200, message: "OK" });
                    } else {
                        res.status(400).send({ code: 400, message: "Can't find conversation ID" });
                    }
                }
            } else {
                res.status(401).send({ code: 401, message: "Not Authorized" });
            }
        });
    } else {
        throw new Error("Microsoft Bot Builder Direct Line Secret was not provided in options or the environment variable MICROSOFT_DIRECTLINE_SECRET");
    }
};

// this method is to trigger the handoff (useful for when you want a luis dialog to trigger the handoff, instead of the keyword)
async function triggerHandoff(bot, connector, session) {
    const message = session.message;
    const conversation = await handoff.getConversation({ customerConversationId: message.address.conversation.id }, message.address);
    if (conversation.state === ConversationState.Bot) {
        // do not log this to prevent duplicates
        // await handoff.addToTranscript({ customerConversationId: conversation.customer.conversation.id }, message);
        await handoff.queueCustomerForAgent({ customerConversationId: conversation.customer.conversation.id });
        // send notification of a new help request in support
        const msg = new builder.Message().address(support_address as any);

        // if is member of team, also mention it
        // tslint:disable-next-line: variable-name
        let team_text = session.message.address.user.name;
        if ((session.message as any).channelId === "msteams" && message.address.conversation.isGroup) {
            // get the team name
            const teamName = await new Promise((resolve, reject) => {
                connector.fetchTeamInfo(( session.message.address as builder.IChatConnectorAddress).serviceUrl,
                                              session.message.sourceEvent.team.id || null, (err, result) => {
                    if (err) { reject(err); } else { resolve(result.name); }
                });
            });
            team_text += " from " + teamName;
        }
        team_text += " needs help. Last message:\n" + message.text;

        const herocard = new builder.HeroCard(session)
        .text(team_text)
        .buttons([
            builder.CardAction.imBack(session, "connect " + message.address.conversation.id, "Connect"),
            builder.CardAction.imBack(session, "watch " + message.address.conversation.id, "Watch"),
            builder.CardAction.imBack(session, "history " + message.address.conversation.id, "Chat logs"),
        ]);
        // attach the card to the reply message
        msg.addAttachment(herocard);
        bot.send(msg);
        // commented out because we don't want the user to know
        // session.endConversation("Connecting you to the next available agent.");
        return;
    }
}

module.exports = { setup, triggerHandoff };

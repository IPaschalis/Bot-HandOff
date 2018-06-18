"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const builder = require("botbuilder");
const mongoose_provider_1 = require("./mongoose-provider");
const teams = require("botbuilder-teams");
const indexExports = require('./index');
// Options for state of a conversation
// Customer talking to bot, waiting for next available agent or talking to an agent
var ConversationState;
(function (ConversationState) {
    ConversationState[ConversationState["Bot"] = 0] = "Bot";
    ConversationState[ConversationState["Waiting"] = 1] = "Waiting";
    ConversationState[ConversationState["Agent"] = 2] = "Agent";
    ConversationState[ConversationState["Watch"] = 3] = "Watch";
})(ConversationState = exports.ConversationState || (exports.ConversationState = {}));
;
class Handoff {
    // if customizing, pass in your own check for isAgent and your own versions of methods in defaultProvider
    constructor(bot, connector, isAgent, provider = new mongoose_provider_1.MongooseProvider()) {
        this.bot = bot;
        this.connector = connector;
        this.isAgent = isAgent;
        this.provider = provider;
        this.connectCustomerToAgent = (by, nextState, agentAddress) => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.connectCustomerToAgent(by, nextState, agentAddress);
        });
        this.connectCustomerToBot = (by) => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.connectCustomerToBot(by);
        });
        this.queueCustomerForAgent = (by) => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.queueCustomerForAgent(by);
        });
        this.addToTranscript = (by, message) => __awaiter(this, void 0, void 0, function* () {
            let from = by.agentConversationId ? 'Agent' : 'Customer';
            return yield this.provider.addToTranscript(by, message, from);
        });
        this.getConversation = (by, customerAddress, teamId, teamName, tenantId) => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.getConversation(by, customerAddress, teamId, teamName, tenantId);
        });
        this.getCurrentConversations = () => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.getCurrentConversations();
        });
        this.getCurrentTeams = () => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.getCurrentTeams();
        });
        this.getTeamConversations = (teamName) => __awaiter(this, void 0, void 0, function* () {
            return yield this.provider.getTeamConversations(teamName);
        });
        this.provider.init();
    }
    routingMiddleware() {
        return {
            botbuilder: (session, next) => {
                // Pass incoming messages to routing method
                if (session.message.type === 'message') {
                    this.routeMessage(session, next);
                }
                else {
                    // allow messages of non 'message' type through 
                    next();
                }
            },
            send: (event, next) => __awaiter(this, void 0, void 0, function* () {
                // Messages sent from the bot do not need to be routed
                // skip agent messages
                if (event.address.conversation.id.split(';')[0] == indexExports._supportChannelId)
                    next();
                // Not all messages from the bot are type message, we only want to record the actual messages  
                else if (event.type === 'message' && !event.entities) {
                    const message = event;
                    const customerConversation = yield this.getConversation({ customerConversationId: event.address.conversation.id });
                    // send message to agent observing conversation
                    if (customerConversation.state === ConversationState.Watch) {
                        this.bot.send(new builder.Message().address(customerConversation.agent).text(message.text));
                    }
                    this.transcribeMessageFromBot(event, next);
                }
                else {
                    //If not a message (text), just send to user without transcribing
                    next();
                }
            })
        };
    }
    routeMessage(session, next) {
        if (this.isAgent(session)) {
            this.routeAgentMessage(session);
        }
        else {
            this.routeCustomerMessage(session, next);
        }
    }
    routeAgentMessage(session) {
        return __awaiter(this, void 0, void 0, function* () {
            const message = session.message;
            const conversation = yield this.getConversation({ agentConversationId: message.address.conversation.id }, message.address);
            yield this.addToTranscript({ agentConversationId: message.address.conversation.id }, message);
            // if the agent is not in conversation, no further routing is necessary
            if (!conversation)
                return;
            //if state of conversation is not 2, don't route agent message
            if (conversation.state !== ConversationState.Agent) {
                // error state -- should not happen
                session.send("Shouldn't be in this state - agent should have been cleared out.");
                return;
            }
            // send text that agent typed to the customer they are in conversation with
            this.bot.send(new builder.Message().address(conversation.customer).text(message.text).addEntity({ "agent": true }));
        });
    }
    routeCustomerMessage(session, next) {
        return __awaiter(this, void 0, void 0, function* () {
            const message = session.message;
            // method will either return existing conversation or a newly created conversation if this is first time we've heard from customer
            let teamId = null;
            let tenantId = null;
            let teamName = null;
            if (session.message.channelId = "msteams") {
                teamId = session.message.sourceEvent.teamsTeamId || null;
                tenantId = teams.TeamsMessage.getTenantId(session.message);
                // if in a team, get the name
                if (message.address.conversation.isGroup) {
                    teamName = yield new Promise((resolve, reject) => {
                        this.connector.fetchTeamInfo(session.message.address.serviceUrl, teamId, (err, result) => {
                            if (err) {
                                reject(null);
                            }
                            else {
                                resolve(result.name);
                            }
                        });
                    });
                }
            }
            const conversation = yield this.getConversation({ customerConversationId: message.address.conversation.id }, message.address, teamId, teamName, tenantId);
            yield this.addToTranscript({ customerConversationId: conversation.customer.conversation.id }, message);
            switch (conversation.state) {
                case ConversationState.Bot:
                    return next();
                case ConversationState.Waiting:
                    return next();
                case ConversationState.Watch:
                    this.bot.send(new builder.Message().address(conversation.agent).text(message.text));
                    return next();
                case ConversationState.Agent:
                    if (!conversation.agent) {
                        session.send("No agent address present while customer in state Agent");
                        console.log("No agent address present while customer in state Agent");
                        return;
                    }
                    this.bot.send(new builder.Message().address(conversation.agent).text(message.text));
                    return;
            }
        });
    }
    // These methods are wrappers around provider which handles data
    transcribeMessageFromBot(message, next) {
        this.provider.addToTranscript({ customerConversationId: message.address.conversation.id }, message, 'Bot');
        next();
    }
    getCustomerTranscript(by, session) {
        return __awaiter(this, void 0, void 0, function* () {
            const customerConversation = yield this.getConversation(by);
            if (customerConversation) {
                let text = '';
                // only return the last 10
                customerConversation.transcript.slice(-10).forEach(transcriptLine => text += `**${transcriptLine.from}** (*${new Date(transcriptLine.timestamp).toLocaleString()} UTC*): ${transcriptLine.text}\n\n`);
                session.send(text);
            }
            else {
                session.send('No Transcript to show. Try entering a username or try again when connected to a customer');
            }
        });
    }
}
exports.Handoff = Handoff;
;
//# sourceMappingURL=handoff.js.map
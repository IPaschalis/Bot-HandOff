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
const teams = require("botbuilder-teams");
const handoff_1 = require("./handoff");
const indexExports = require("./index");
function commandsMiddleware(bot, handoff) {
    return {
        botbuilder: (session, next) => {
            if (session.message.type === "message") {
                command(session, next, handoff, bot);
            }
            else {
                // allow messages of non 'message' type through
                next();
            }
        },
    };
}
exports.commandsMiddleware = commandsMiddleware;
function command(session, next, handoff, bot) {
    if (handoff.isAgent(session)) {
        agentCommand(session, next, handoff, bot);
    }
    else {
        customerCommand(session, next, handoff, bot);
    }
}
function agentCommand(session, next, handoff, bot) {
    return __awaiter(this, void 0, void 0, function* () {
        const message = session.message;
        const conversation = yield handoff.getConversation({ agentConversationId: message.address.conversation.id });
        let inputWords = message.text.split(/\s+/);
        if (inputWords.length === 0) {
            return;
        }
        // capture the adaptive card content and handle it here (for list conversations)
        if (message && message.value) {
            inputWords = message.value.action.split(/\s+/);
        }
        // Commands to execute whether connected to a customer or not
        switch (inputWords[0]) {
            case "options":
                sendAgentCommandOptions(session);
                return;
            case "list":
                if (inputWords.length === 1) {
                    yield currentTeams(session, handoff);
                }
                else {
                    // show the conversation of that Team
                    const team = inputWords.slice(1).join(" ");
                    const conversations = (yield handoff.getTeamConversations(team));
                    if (!conversations) {
                        session.send(`Team '${team}' does not exist. Connect to one of the Teams below.`);
                        yield currentTeams(session, handoff);
                    }
                    else {
                        yield currentConversations(session, handoff, conversations);
                    }
                }
                return;
            case "history":
                yield handoff.getCustomerTranscript(inputWords.length > 1
                    ? { customerConversationId: inputWords.slice(1).join(" ") }
                    : { agentConversationId: message.address.conversation.id }, session);
                return;
            case "waiting":
                if (conversation) {
                    // disconnect from current conversation if already talking
                    disconnectCustomer(conversation, handoff, session, bot);
                }
                const waitingConversation = yield handoff.connectCustomerToAgent({ bestChoice: true }, handoff_1.ConversationState.Agent, message.address);
                if (waitingConversation) {
                    const team = yield handoff.getConversationTeam(waitingConversation._id);
                    session.send(`You are connected to **${waitingConversation.customer.user.name}** (${waitingConversation.customer.user.id}) from **${team.teamName}**`);
                }
                else {
                    session.send("No customers waiting.");
                }
                return;
            case "connect":
            case "watch":
                // const newConversation = await handoff.connectCustomerToAgent(
                //    inputWords.length > 1
                //        ? { customerId: inputWords.slice(1).join(' ') }
                //        : { bestChoice: true },
                //    ConversationState.Agent,
                //    message.address
                // );
                let newConversation;
                if (inputWords[0] === "connect") {
                    newConversation = yield handoff.connectCustomerToAgent(inputWords.length > 1
                        ? { customerConversationId: inputWords.slice(1).join(" ") }
                        : { bestChoice: true }, handoff_1.ConversationState.Agent, message.address);
                }
                else {
                    newConversation = yield handoff.connectCustomerToAgent({ customerConversationId: inputWords.slice(1).join(" ") }, handoff_1.ConversationState.Watch, message.address);
                }
                if (newConversation) {
                    const team = yield handoff.getConversationTeam(newConversation._id);
                    session.send(`You are connected to **${newConversation.customer.user.name}** (${newConversation.customer.user.id}) from **${team.teamName}**`);
                }
                else {
                    session.send("No customers waiting.");
                }
                if (message.text === "disconnect") {
                    disconnectCustomer(conversation, handoff, session, bot);
                }
                return;
            case "disconnect":
                disconnectCustomer(conversation, handoff, session, bot);
                return;
            default:
                if (conversation && conversation.state === handoff_1.ConversationState.Agent) {
                    return next();
                }
                sendAgentCommandOptions(session);
                return;
        }
    });
}
// TODO: fix undefined "this" error
function customerCommand(session, next, handoff, bot) {
    return __awaiter(this, void 0, void 0, function* () {
        const message = session.message;
        const customerStartHandoffCommandRegex = new RegExp("^" + indexExports._customerStartHandoffCommand + "$", "gi");
        if (customerStartHandoffCommandRegex.test(message.text)) {
            // lookup the conversation (create it if one doesn't already exist)
            // also pass the teamId
            let teamId = null;
            let teamName = null;
            let tenantId = null;
            if (session.message.address.channelId === "msteams") {
                teamId = session.message.sourceEvent.teamsTeamId;
                tenantId = teams.TeamsMessage.getTenantId(session.message);
                teamName = yield new Promise((resolve, reject) => {
                    this.connector.fetchTeamInfo(session.message.address.serviceUrl, session.message.sourceEvent.team.id || null, (err, result) => {
                        if (err) {
                            reject(err);
                        }
                        else {
                            resolve(result.name);
                        }
                    });
                });
            }
            const conversation = yield handoff.getConversation({ customerConversationId: message.address.conversation.id }, message.address, teamId, teamName, tenantId);
            if (conversation.state === handoff_1.ConversationState.Bot) {
                // send notification of a new help request in support
                const reply = new teams.TeamsMessage();
                reply.address(indexExports.support_address);
                // if is member of team, also mention it
                let teamText = "";
                if (session.message.channelId === "msteams" && message.address.conversation.isGroup) {
                    teamText = " from " + teamName;
                }
                reply.text(session.message.address.user.name + teamText + " needs help.");
                // temporarily disable to prevent spam TODO: enable again
                // bot.send(reply);
                yield handoff.addToTranscript({ customerConversationId: conversation.customer.conversation.id }, message);
                yield handoff.queueCustomerForAgent({ customerConversationId: conversation.customer.conversation.id });
                // endConversation not supported in Teams
                session.send("Connecting you to the next available agent.");
                return;
            }
        }
        return next();
    });
}
function sendAgentCommandOptions(session) {
    const commands = "### Agent Options \n\n - Type *waiting* to connect to customer who has been waiting longest.\n\n - Type *connect { user id }* to connect to a specific conversation\n\n - Type *watch { user id }* to monitor a customer conversation\n\n - Type *history { user id }* to see a transcript of a given user\n\n - Type *list* to see a list of all current conversations.\n\n - Type *disconnect* while talking to a user to end a conversation.\n\n - Type *options* at any time to see these options again.";
    session.send(commands);
    return;
}
function currentConversations(session, handoff, conversations) {
    return __awaiter(this, void 0, void 0, function* () {
        // if we didn't pass the conversations parameters, find all conversations
        if (!conversations) {
            conversations = yield handoff.getCurrentConversations();
        }
        if (conversations.length === 0) {
            return "No customers are in conversation.";
        }
        // keep the last 10 (mst recent) to avoid errors about too big card
        conversations = conversations.slice(-10).reverse();
        const items = [];
        conversations.forEach((conversation) => {
            const stateText = {
                Bot: "talking to bot",
                Agent: "talking to agent",
                Waiting: "waiting for agent",
                Watch: "monitored by agent",
            };
            const item = {
                type: "ColumnSet",
                separator: true,
                columns: [
                    {
                        type: "Column",
                        separator: true,
                        width: "stretch",
                        items: [{
                                type: "TextBlock",
                                weight: "bolder",
                                text: conversation.customer.user.name,
                                wrap: true,
                            },
                            {
                                type: "TextBlock",
                                spacing: "none",
                                text: stateText[handoff_1.ConversationState[conversation.state]],
                            },
                            {
                                type: "TextBlock",
                                spacing: "none",
                                text: `last msg: ${new Date(conversation.transcript[conversation.transcript.length - 1].timestamp).toLocaleString()}`,
                                wrap: true,
                            }],
                    },
                    {
                        type: "Column",
                        width: "auto",
                        items: [{
                                color: "accent",
                                type: "TextBlock",
                                text: "[CONNECT]",
                                weight: "bolder",
                            }],
                        selectAction: {
                            type: "Action.Submit",
                            title: "Action.Submit",
                            data: {
                                action: `connect ${conversation.customer.conversation.id}`,
                            },
                        },
                    },
                    {
                        type: "Column",
                        width: "auto",
                        items: [{
                                color: "accent",
                                type: "TextBlock",
                                text: "[WATCH]",
                                weight: "bolder",
                            }],
                        selectAction: {
                            type: "Action.Submit",
                            title: "Action.Submit",
                            data: {
                                action: `watch ${conversation.customer.conversation.id}`,
                            },
                        },
                    },
                    {
                        type: "Column",
                        width: "auto",
                        items: [{
                                color: "accent",
                                type: "TextBlock",
                                text: "[HISTORY]",
                                weight: "bolder",
                            }],
                        selectAction: {
                            type: "Action.Submit",
                            title: "Action.Submit",
                            data: {
                                action: `history ${conversation.customer.conversation.id}`,
                            },
                        },
                    },
                ],
            };
            items.push(item);
        });
        const attachment = {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                type: "AdaptiveCard",
                version: "1.0",
                body: [{
                        type: "Container",
                        items,
                    }],
            },
        };
        const msg = new builder.Message(session)
            .addAttachment(attachment);
        session.send(msg);
        // let text = '### Current Conversations \n';
        // text += "Please use the conversation's ID to connect.\n\n";
        // conversations.forEach(conversation => {
        //     const starterText = ` - **${conversation.customer.user.name}** *(convID: ${conversation.customer.conversation.id})*`;
        //     switch (ConversationState[conversation.state]) {
        //         case 'Bot':
        //             text += starterText + ' is talking to the bot\n';
        //             break;
        //         case 'Agent':
        //             text += starterText + ' is talking to an agent\n';
        //             break;
        //         case 'Waiting':
        //             text += starterText + ' is waiting to talk to an agent\n';
        //             break;
        //         case 'Watch':
        //             text += starterText + ' is being monitored by an agent\n';
        //             break
        //     }
        //     text += `| **last msg:** ${new Date(conversation.transcript[conversation.transcript.length-1].timestamp).toLocaleString()}\n`;
        // });
        // return text;
    });
}
function currentTeams(session, handoff) {
    return __awaiter(this, void 0, void 0, function* () {
        const allTeams = yield handoff.getCurrentTeams();
        if (allTeams.length === 0) {
            session.send("No customers are in conversation.");
            return;
        }
        const msg = new builder.Message();
        const buttons = [];
        allTeams.forEach((team) => {
            buttons.push(builder.CardAction.imBack(session, "list " + team.teamName, team.teamName));
        });
        const herocard = new builder.HeroCard(session)
            .text("Current Teams \n")
            .buttons(buttons);
        msg.addAttachment(herocard);
        session.send(msg);
    });
}
function disconnectCustomer(conversation, handoff, session, bot) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const isconnected = yield handoff.connectCustomerToBot({ customerConversationId: conversation.customer.conversation.id });
            // Send message to agent
            if (isconnected) {
                session.send(`Customer ${conversation.customer.user.name} (${conversation.customer.user.id}) is now connected to the bot.`);
            }
            // do not inform customer of agent disconnect now
            // if (bot && conversation.state!=ConversationState.Watch) {
            //    //Send message to customer
            //    var reply = new builder.Message()
            //        .address(conversation.customer)
            //        .text('Agent has disconnected, you are now speaking to the bot.');
            //    bot.send(reply);
            // }
        }
        catch (err) {
            if (err.message === "Cannot read property 'customer' of null") {
                session.send("No customer is connected to the bot.");
            }
            else {
                session.send(err.message);
            }
        }
    });
}
//# sourceMappingURL=commands.js.map
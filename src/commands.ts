import * as builder from 'botbuilder';
import * as teams from 'botbuilder-teams'
import { Conversation, ConversationState, Handoff } from './handoff';
const indexExports = require('./index');

export function commandsMiddleware(bot: builder.UniversalBot, handoff: Handoff) {
    return {
        botbuilder: (session: builder.Session, next: Function) => {
            if (session.message.type === 'message') {
                command(session, next, handoff, bot);
            } else {
                // allow messages of non 'message' type through 
                next();
            }
        }
    }
}

function command(session: builder.Session, next: Function, handoff: Handoff, bot: builder.UniversalBot) {
    if (handoff.isAgent(session)) {
        agentCommand(session, next, handoff, bot);
    } else {
        customerCommand(session, next, handoff, bot);
    }
}

async function agentCommand(
    session: builder.Session,
    next: Function,
    handoff: Handoff,
    bot: builder.UniversalBot
) {

    const message = session.message;
    const conversation = await handoff.getConversation({ agentConversationId: message.address.conversation.id });
    const inputWords = message.text.split(/\s+/);

    if (inputWords.length == 0)
        return;

    // Commands to execute whether connected to a customer or not
    switch (inputWords[0]) {
        case 'options':
            sendAgentCommandOptions(session);
            return;
        case 'list':
            session.send(await currentConversations(handoff));
            return;
        case 'history':
            await handoff.getCustomerTranscript(
                inputWords.length > 1
                    ? { customerId: inputWords.slice(1).join(' ') }
                    : { agentConversationId: message.address.conversation.id },
                session);
            return;
        case 'waiting':
            if (conversation) {
                //disconnect from current conversation if already talking
                disconnectCustomer(conversation, handoff, session, bot);
            }
            const waitingConversation = await handoff.connectCustomerToAgent(
                { bestChoice: true },
                ConversationState.Agent,
                message.address
            );
            if (waitingConversation) {
                session.send(`You are connected to ${waitingConversation.customer.user.name} (${waitingConversation.customer.user.id})`);
            } else {
                session.send("No customers waiting.");
            }
            return;
        case 'connect':
        case 'watch':
            //const newConversation = await handoff.connectCustomerToAgent(
            //    inputWords.length > 1
            //        ? { customerId: inputWords.slice(1).join(' ') }
            //        : { bestChoice: true },
            //    ConversationState.Agent,
            //    message.address
            //);
            let newConversation;
            if(inputWords[0] == 'connect') {
                newConversation = await handoff.connectCustomerToAgent(
                    inputWords.length > 1
                        ? { customerConversationId: inputWords.slice(1).join(' ') }
                        : { bestChoice: true },
                    ConversationState.Agent,
                    message.address
                );
            }
            else
                newConversation = await handoff.connectCustomerToAgent({ customerConversationId: inputWords.slice(1).join(' ') }, ConversationState.Watch, message.address);

            if (newConversation) {
                session.send(`You are connected to ${newConversation.customer.user.name} (${newConversation.customer.user.id})`);
            }
            else {
                session.send("No customers waiting.");
            }

            if (message.text === 'disconnect') {
                disconnectCustomer(conversation, handoff, session, bot);
            }
            return;
        case 'disconnect':
            disconnectCustomer(conversation, handoff, session, bot);
            return;
        default:
            if (conversation && conversation.state === ConversationState.Agent) {
                return next();
            }
            sendAgentCommandOptions(session);
            return;
    }
}

async function customerCommand(session: builder.Session, next: Function, handoff: Handoff, bot:builder.UniversalBot) {
    const message = session.message;
    const customerStartHandoffCommandRegex = new RegExp("^" + indexExports._customerStartHandoffCommand + "$", "gi");
    if (customerStartHandoffCommandRegex.test(message.text)) {
        // lookup the conversation (create it if one doesn't already exist)

        //also pass the teamId
        let teamId = null;
        if ((session.message as any).channelId == "msteams") {
            teamId = session.message.sourceEvent.teamsTeamId;
        }

        const conversation = await handoff.getConversation({ customerConversationId: message.address.conversation.id }, message.address, teamId);
        if (conversation.state == ConversationState.Bot) {

            //send notification of a new help request in support 
            var reply = new teams.TeamsMessage()
            reply.address(indexExports.support_address);

            //if is member of team, also mention it
            let team_text = '';
            if ((session.message as any).channelId == 'msteams' && message.address.conversation.isGroup) {
                team_text = ' from ' + session.message.sourceEvent.teamsTeamId;
            }

            reply.text(session.message.address.user.name + team_text + ' needs help.');
            bot.send(reply);

            await handoff.addToTranscript({ customerConversationId: conversation.customer.conversation.id }, message);
            await handoff.queueCustomerForAgent({ customerConversationId: conversation.customer.conversation.id });
            // endConversation not supported in Teams 
            session.send("Connecting you to the next available agent.");
            return;
        }
    }
    return next();
}

function sendAgentCommandOptions(session: builder.Session) {
    const commands = ' ### Agent Options\n - Type *waiting* to connect to customer who has been waiting longest.\n - Type *connect { user id }* to connect to a specific conversation\n - Type *watch { user id }* to monitor a customer conversation\n - Type *history { user id }* to see a transcript of a given user\n - Type *list* to see a list of all current conversations.\n - Type *disconnect* while talking to a user to end a conversation.\n - Type *options* at any time to see these options again.';
    session.send(commands);
    return;
}

async function currentConversations(handoff: Handoff): Promise<string> {
    const conversations = await handoff.getCurrentConversations();
    if (conversations.length === 0) {
        return "No customers are in conversation.";
    }

    let text = '### Current Conversations \n';
    text += "Please use the user's ID to connect with them.\n\n";
    conversations.forEach(conversation => {
        const starterText = ` - **${conversation.customer.user.name}** *(convID: ${conversation.customer.conversation.id})*`;
        switch (ConversationState[conversation.state]) {
            case 'Bot':
                text += starterText + ' is talking to the bot\n';
                break;
            case 'Agent':
                text += starterText + ' is talking to an agent\n';
                break;
            case 'Waiting':
                text += starterText + ' is waiting to talk to an agent\n';
                break;
            case 'Watch':
                text += starterText + ' is being monitored by an agent\n';
                break
        }
        text += `| **last msg:** ${new Date(conversation.transcript[conversation.transcript.length-1].timestamp).toLocaleString()}\n`;
    });

    return text;
}

async function disconnectCustomer(conversation: Conversation, handoff: any, session: builder.Session, bot?: builder.UniversalBot) {
    if (await handoff.connectCustomerToBot({ customerConversationId: conversation.customer.conversation.id })) {
        //Send message to agent
        session.send(`Customer ${conversation.customer.user.name} (${conversation.customer.user.id}) is now connected to the bot.`);

		// do not inform customer if agent is watching
        if (bot && conversation.state!=ConversationState.Watch) {
            //Send message to customer
            var reply = new builder.Message()
                .address(conversation.customer)
                .text('Agent has disconnected, you are now speaking to the bot.');
            bot.send(reply);
        }
    }
}
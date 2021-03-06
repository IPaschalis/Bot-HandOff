// tslint:disable: interface-name
// tslint:disable: variable-name
import * as builder from "botbuilder";
import * as bluebird from "bluebird";
import * as request from "request";
import * as _ from "lodash";
import mongoose = require("mongoose");
mongoose.Promise = bluebird;
// fix deprecation warning
mongoose.set("useFindAndModify", false);

import { By, Conversation, Team, Provider, ConversationState } from "./handoff";

const indexExports = require("./index");

// -------------------
// Bot Framework types
// -------------------
export const IIdentitySchema = new mongoose.Schema({
    id: { type: String, required: true },
    isGroup: { type: Boolean, required: false },
    name: { type: String, required: false },
}, {
        _id: false,
        strict: false,
    });

export const IAddressSchema = new mongoose.Schema({
    bot: { type: IIdentitySchema, required: true },
    channelId: { type: String, required: true },
    conversation: { type: IIdentitySchema, required: false },
    user: { type: IIdentitySchema, required: true },
    id: { type: String, required: false },
    serviceUrl: { type: String, required: false },
    useAuth: { type: Boolean, required: false },
}, {
        strict: false,
        id: false,
        _id: false,
    });

// -------------
// Handoff types
// -------------
export const TranscriptLineSchema = new mongoose.Schema({
    timestamp: {},
    from: String,
    sentimentScore: Number,
    state: Number,
    text: String,
});

export const ConversationSchema = new mongoose.Schema({
    customer: { type: IAddressSchema, required: true },
    agent: { type: IAddressSchema, required: false },
    state: {
        type: Number,
        required: true,
        min: 0,
        max: 3,
    },
    transcript: [TranscriptLineSchema],
});
export interface ConversationDocument extends Conversation, mongoose.Document { }
export const ConversationModel = mongoose.model<ConversationDocument>("Conversation", ConversationSchema);

// Teams collection. It will point to its Conversation IDs
export const TeamSchema = new mongoose.Schema({
    teamId: {type: String, required: false},
    tenantId: {type: String, required: false},
    channel: {type: String, required: true, default: "Teams"},
    teamName: {type: String, required: true},
    conversation: [{
        type: mongoose.Schema.Types.ObjectId,
        ref: "Conversation",
    }],
});
export interface TeamDocument extends Team, mongoose.Document { }
export const TeamModel = mongoose.model<TeamDocument>("Team", TeamSchema);

export const BySchema = new mongoose.Schema({
    bestChoice: Boolean,
    agentConversationId: String,
    customerConversationId: String,
    customerName: String,
    customerId: String,
});
export interface ByDocument extends By, mongoose.Document { }
export const ByModel = mongoose.model<ByDocument>("By", BySchema);
export { mongoose };

// -----------------
// Mongoose Provider
// -----------------
export class MongooseProvider implements Provider {
    public init(): void { }

    public async addToTranscript(by: By, message: builder.IMessage, from: string): Promise<boolean> {
        let sentimentScore = -1;
        const text = message.text;
        let datetime = new Date().toISOString();
        const conversation: Conversation = await this.getConversation(by);

        if (!conversation) { return false; }

        if (from === "Customer") {
            if (indexExports._textAnalyticsKey) { sentimentScore = await this.collectSentiment(text); }
            datetime = message.localTimestamp ? message.localTimestamp : message.timestamp;
        }

        conversation.transcript.push({
            timestamp: datetime,
            from,
            sentimentScore,
            state: conversation.state,
            text,
        });

        if (indexExports._appInsights) {
            // You can't log embedded json objects in application insights, so we are flattening the object to one item.
            // Also, have to stringify the object so functions from mongodb don't get logged
            const latestTranscriptItem = conversation.transcript.length - 1;
            const x = JSON.parse(JSON.stringify(conversation.transcript[latestTranscriptItem]));
            x.botId = conversation.customer.bot.id;
            x.customerId = conversation.customer.user.id;
            x.customerName = conversation.customer.user.name;
            x.customerChannelId = conversation.customer.channelId;
            x.customerConversationId = conversation.customer.conversation.id;
            if (conversation.agent) {
                x.agentId = conversation.agent.user.id;
                x.agentName = conversation.agent.user.name;
                x.agentChannelId = conversation.agent.channelId;
                x.agentConversationId = conversation.agent.conversation.id;
            }
            indexExports._appInsights.client.trackEvent("Transcript", x);
        }

        return await this.updateConversation(conversation);
    }

    public async connectCustomerToAgent(by: By, stateUpdate: ConversationState, agentAddress: builder.IAddress): Promise<Conversation> {
        const conversation: Conversation = await this.getConversation(by);
        if (conversation) {
            conversation.state = stateUpdate;
            conversation.agent = agentAddress;
        }
        const success = await this.updateConversation(conversation);
        if (success) {
            return conversation;
        } else {
            return null;
        }
    }

    public async queueCustomerForAgent(by: By): Promise<boolean> {
        const conversation: Conversation = await this.getConversation(by);
        if (!conversation) {
            return false;
        } else {
            conversation.state = ConversationState.Waiting;
            return await this.updateConversation(conversation);
        }
    }

    public async connectCustomerToBot(by: By): Promise<boolean> {
        const conversation: Conversation = await this.getConversation(by);
        if (!conversation) {
            return false;
        } else {
            conversation.state = ConversationState.Bot;
            if (indexExports._retainData === "true") {
                // if retain data is true, AND the user has spoken to an agent - delete the agent record
                // this is necessary to avoid a bug where the agent cannot connect to another user after disconnecting with a user
                if (conversation.agent) {
                    conversation.agent = null;
                    return await this.updateConversation(conversation);
                } else {
                    // otherwise, just update the conversation
                    return await this.updateConversation(conversation);
                }
            } else {
                // if retain data is false, delete the whole conversation after talking to agent
                if (conversation.agent) {
                    return await this.deleteConversation(conversation);
                } else {
                    // otherwise, just update the conversation
                    return await this.updateConversation(conversation);
                }
            }
        }
    }

    public async getConversation(by: By, customerAddress?: builder.IAddress, teamId?: string, teamName?: string, tenantId?: string): Promise<Conversation> {
        if (by.customerName) {
            const conversation = await ConversationModel.findOne({ "customer.user.name": by.customerName });
            return conversation;
        } else if (by.customerId) {
            const conversation = await ConversationModel.findOne({ "customer.user.id": by.customerId });
            return conversation;
        } else if (by.agentConversationId) {
            const conversation = await ConversationModel.findOne({ "agent.conversation.id": by.agentConversationId });
            if (conversation) { return conversation; } else { return null; }
        } else if (by.customerConversationId) {
            let conversation: Conversation = await ConversationModel.findOne({ "customer.conversation.id": by.customerConversationId });
            if (!conversation && customerAddress) {
                conversation = await this.createConversation(customerAddress, teamId, teamName, tenantId);
            }
            return conversation;
        } else if (by.bestChoice) {
            let waitingLongest = await this.getCurrentConversations();
            waitingLongest = waitingLongest
                .filter((conversation) => conversation.state === ConversationState.Waiting)
                .sort((x, y) => y.transcript[y.transcript.length - 1].timestamp - x.transcript[x.transcript.length - 1].timestamp);
            return waitingLongest.length > 0 && waitingLongest[0];
        }
        return null;
    }

    public async getCurrentConversations(): Promise<Conversation[]> {
        let conversations;
        try {
            conversations = await ConversationModel.find();
        } catch (error) {
            console.log("Failed loading conversations");
            console.log(error);
        }
        return conversations;
    }

    public async getTeamConversations(teamName: string): Promise<Conversation[]> {
        let conversations;
        try {
            // find the coresponding conversations from the ids
            const model = await TeamModel.findOne({teamName}).select("conversation").populate("conversation");
            conversations = model.conversation;
        } catch (error) {
            console.log("Failed loading conversations");
            console.log(error);
        }
        return conversations;
    }

    public async getConversationTeam(convId: string): Promise<Team> {
        let team;
        const convIdObj = mongoose.Types.ObjectId(convId);
        try {
            team = await TeamModel.findOne({conversation: convIdObj});
        } catch (error) {
            console.log("Failed getting conversation's team");
            console.log(error);
        }
        return team;
    }

    public async getCurrentTeams(): Promise<Team[]> {
        let teams;
        try {
            teams = await TeamModel.find();
        } catch (error) {
            console.log("Failed loading Teams");
            console.log(error);
        }
        return teams;
    }

    private async createConversation(customerAddress: builder.IAddress, teamId: string, teamName: string,  tenantId: string): Promise<Conversation> {
        const conversation = await ConversationModel.create({
            customer: customerAddress,
            state: ConversationState.Bot,
            transcript: [],
        });

        // find the team this conversation belongs to
        if (teamId == null) {
            teamId = null;
            teamName = "Personal Chat";
        }
        let team = await TeamModel.findOne({teamName});
        // if it doesn't exist create it
        if (!team) {
            team = await this.createTeam(teamId, teamName, tenantId);
        }
        // add the conversation to the team
        const success = await this.updateTeam(team, conversation._id);
        if (!success) { return null; }

        return conversation;
    }

    private async createTeam(teamId: string, teamName: string, tenantId: string): Promise<TeamDocument> {
        return await TeamModel.create({
            teamId,
            teamName,
            tenantId,
            conversation: [],
        });
    }

    private async updateTeam(team: TeamDocument, convid: string): Promise<boolean> {
        team.conversation.push(convid);
        return new Promise<boolean>((resolve, reject) => {
            TeamModel.findByIdAndUpdate((team as any)._id, team).then((error) => {
                resolve(true);
            }).catch((error) => {
                console.log("Failed to update team");
                console.log(team);
                resolve(false);
            });
        });
    }

    private async updateConversation(conversation: Conversation): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            ConversationModel.findByIdAndUpdate((conversation as any)._id, conversation).then((error) => {
                resolve(true);
            }).catch((error) => {
                console.log("Failed to update conversation");
                console.log(conversation as any);
                resolve(false);
            });
        });
    }

    private async deleteConversation(conversation: Conversation): Promise<boolean> {
        return new Promise<boolean>((resolve) => {
            ConversationModel.findByIdAndRemove((conversation as any)._id).then((error) => {
                resolve(true);
            });
        });
    }

    private async collectSentiment(text: string): Promise<number> {
        if (text == null || text === "") { return; }
        const _sentimentUrl = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment";
        const _sentimentId = "bot-analytics";
        const _sentimentKey = indexExports._textAnalyticsKey;

        const options = {
            url: _sentimentUrl,
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Ocp-Apim-Subscription-Key": _sentimentKey,
            },
            json: true,
            body: {
                documents: [
                    {
                        language: "en",
                        id: _sentimentId,
                        text,
                    },
                ],
            },
        };

        return new Promise<number>((resolve, reject) => {
            request(options, (error, response, body) => {
                if (error) { reject(error); }
                const result: any = _.find(body.documents, { id: _sentimentId }) || {};
                const score = result.score || null;
                resolve(score);
            });
        });
    }
}

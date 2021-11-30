import { BotDeclaration, PreventIframe, BotCallingWebhook } from "express-msteams-host";
import * as debug from "debug";
import { CardFactory, ConversationState, MemoryStorage, UserState, TurnContext, BotFrameworkAdapter, ConversationReference, AdaptiveCardInvokeValue, AdaptiveCardInvokeResponse, SigninStateVerificationQuery } from "botbuilder";
import { DialogBot } from "./dialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import WelcomeCard from "./cards/welcomeCard";
import callingHandler from "./callingHandler";
import express = require("express");
import chalk = require("chalk");
import NewMeetingCard from "./cards/newMeetingCard";
import meetings from "../meetingManager";
import * as incidents from "../incidents.json";
import { stripHtml } from "string-strip-html";
import * as fs from "fs";
import * as path from "path";
import * as ffmpeg from "fluent-ffmpeg";
import * as cog from "microsoft-cognitiveservices-speech-sdk";

import { JsonDB } from "node-json-db";
import { Config } from "node-json-db/dist/lib/JsonDBConfig";
import AfterMeetingCard from "./cards/afterMeetingCard";
import { SsoOAuthHelper } from "./SsoAuthHelper";
import { Client } from "@microsoft/microsoft-graph-client";
import { ChatMessage } from "@microsoft/microsoft-graph-types-beta";
import { AzureKeyCredential, TextAnalyticsClient } from "@azure/ai-text-analytics";
import axios from "axios";

export const conversationReferences = new JsonDB(new Config("conversationRefs.db.json", true, false, "/"));

// Initialize debug logging module
const log = debug("msteams");

// store the adapter globally
let _adapter: BotFrameworkAdapter;

/**
 * Implementation for meetai Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_ID,
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_PASSWORD)
@PreventIframe("/meetaiBot/aboutMeetaiBot.html")
export class MeetaiBot extends DialogBot {
    public _ssoOAuthHelper: SsoOAuthHelper;

    constructor(conversationState: ConversationState, userState: UserState, private adapter: BotFrameworkAdapter) {
        super(conversationState, userState, new MainDialog());
        _adapter = adapter;
        this._ssoOAuthHelper = new SsoOAuthHelper(process.env.SSO_CONNECTION_NAME as string, conversationState);
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            if (membersAdded && context.activity.conversation && context.activity.conversation.id) {
                // This is a meeting
                const id = context.activity.conversation.id;
                const meeting = meetings.getByThreadId(id);
                if (meeting) {
                    log(chalk.greenBright("Existing meeting"));
                    await this.sendNewMeetingCard(context, incidents.find(i => i.id === meeting.data.incident));
                } else {
                    log(chalk.greenBright("Unknown meeting"));
                }
            } else if (membersAdded && membersAdded.length > 0) {
                for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                    if (membersAdded[cnt].id !== context.activity.recipient.id) {
                        await this.sendWelcomeCard(context);
                    }
                }
            }
            await next();
        });

        this.onConversationUpdate(async (context, next) => {
            if (context.activity && context.activity) {
                if (context.activity.conversation.conversationType === "personal") {
                    if (!conversationReferences.exists(`/${context.activity.from.aadObjectId}`)) {
                        conversationReferences.push(`/${context.activity.from.aadObjectId}`, TurnContext.getConversationReference(context.activity));
                    }
                }
            }
            await next();
        });
    }

    // https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/universal-action-model
    public async onAdaptiveCardInvoke(context: TurnContext, invokeValue: AdaptiveCardInvokeValue): Promise<AdaptiveCardInvokeResponse> {
        if (invokeValue.action.verb === "generateSummary") {
            log("Generate summary");
            // https://blog.aterentiev.com/ms-graph-get-driveitem-by-file-absolute
            // https://graph.microsoft.com/beta/shares/u!aHR0cHM6Ly93aWN0b3JkZXYtbXkuc2hhcmVwb2ludC5jb20vOnY6L2cvcGVyc29uYWwvd193aWxlbl93aWN0b3JkZXZfb25taWNyb3NvZnRfY29tL0VZRnllSkVfRWtOQ2tpeUlTTkNTbzljQnFac014WmJ6LUdRcXpGZ0ZacDZFSlE_ZW1haWw9dy53aWxlbiU0MHdpY3RvcmRldi5vbm1pY3Jvc29mdC5jb20/driveItem
            // .downloadUrl -> AZ Cog svcs
            // https://docs.microsoft.com/en-us/azure/media-services/latest/analyze-videos-tutorial
            // --> https://docs.microsoft.com/en-us/azure/media-services/latest/job-input-from-http-how-to
            try {
                const token = await this.adapter.getAadTokens(context, "AzureAD", ["https://graph.microsoft.com"]);
                const client = Client.init({
                    authProvider: (cb) => {
                        cb(null, token["https://graph.microsoft.com"].token);
                    }
                });
                const meeting = meetings.getById(invokeValue.action.data.id as string);
                if (meeting === undefined) {
                    return {
                        statusCode: 200,
                        type: "application/vnd.microsoft.activity.message",
                        value: "Meeting not found!"
                    } as any;
                }

                // store the conversation reference due to callbacks etc
                const ref = conversationReferences.getData(`/${context.activity.from.aadObjectId}`) as Partial<ConversationReference>;

                // get all the messages (TODO: implement paging)
                const response = await client.api(`/chats/${meeting.threadId}/messages`).version("beta").get();
                const messages = response.value as ChatMessage[];

                // get all chat messages
                const texts = messages
                    .filter(m => m.messageType === "message" && m.body && m.body.content)
                    .map(m => stripHtml(m.body!.content as string).result)
                    .filter(t => t.length !== 0)
                    .reverse();

                log(chalk.yellow(texts));
                if (texts.length > 0) {
                    // pass to cognitive services
                    const textClient = new TextAnalyticsClient(process.env.AZ_COG_ENDPOINT as string, new AzureKeyCredential(process.env.AZ_COG_KEY as string));
                    textClient.analyzeSentiment(texts).then(results => {
                        // log(results);
                        let positive = 0; let neutral = 0; let negative = 0;
                        results.forEach((r: any) => {
                            if (r.confidenceScores) {
                                positive += r.confidenceScores.positive;
                                negative += r.confidenceScores.negative;
                                neutral += r.confidenceScores.neutral;
                            }
                        });
                        if (positive > negative && positive > neutral) {
                            log("positive");
                            _adapter.continueConversation(ref, async (ctx) => {
                                await ctx.sendActivity("Meeting chat was in general positive 😁");
                            });
                        } else if (neutral > negative) {
                            log("neutral");
                            _adapter.continueConversation(ref, async (ctx) => {
                                await ctx.sendActivity("Meeting chat was in general neutral 😐");
                            });

                        } else {
                            log("negative");
                            _adapter.continueConversation(ref, async (ctx) => {
                                await ctx.sendActivity("Meeting chat was in general negative ☹️");
                            });
                        }
                    });
                }
                // find the transcript
                const transcript = messages
                    .filter(m => m.messageType === "systemEventMessage").pop();
                // TOOD:

                // get the video download url
                const video = messages
                    .filter(m => m.messageType === "systemEventMessage")
                    .filter(m => m.eventDetail &&
                        m.eventDetail["@odata.type"] === "#microsoft.graph.callRecordingEventMessageDetail" &&
                        (m.eventDetail as any).callRecordingStatus === "success")
                    .pop();

                if (video) {
                    const videoDownloadUrl = (video.eventDetail as any).callRecordingUrl;
                    log(videoDownloadUrl);
                    // Create a sharing link
                    // details: https://docs.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0#encoding-sharing-urls
                    const buff = Buffer.from(videoDownloadUrl, "utf-8");
                    const base64 = buff.toString("base64");
                    const encodedUrl = "u!" + base64.replace("/", "_").replace("+", "-").substr(0, base64.length - 1); // remove trailing =

                    // Get the driveItem from Graph
                    client.api(`/shares/${encodedUrl}/driveItem`).version("beta").get().then(r => {
                        log(chalk.cyan(`Anonymous download url: ${r["@microsoft.graph.downloadUrl"]}`));

                        // Download the file locally
                        axios.get(r["@microsoft.graph.downloadUrl"], { responseType: "stream" }).then(blob => {
                            const downloadPath = path.resolve(__dirname, "temp.mp4");
                            const writer = fs.createWriteStream(downloadPath);
                            blob.data.pipe(writer);

                            writer.on("finish", () => {
                                log("file is written");

                                // Use ffmpeg to extract audio
                                const command = ffmpeg();
                                command.on("error", (err) => {
                                    log("An error occurred: " + err.message);
                                });

                                command.on("end", () => {
                                    log("Audio extracted!");

                                    // Use Azure Speach to text to get the transcript
                                    const speechConfig = cog.SpeechConfig.fromSubscription(process.env.AZ_SPEECH_KEY as string, process.env.AZ_SPEECH_REGION as string);
                                    const audioConfig = cog.AudioConfig.fromWavFileInput(fs.readFileSync(path.resolve(__dirname, "audio.wav")));
                                    const recognizer = new cog.SpeechRecognizer(speechConfig, audioConfig);

                                    recognizer.recognizeOnceAsync(async (result) => {
                                        log(`RECOGNIZED: Text=${result.text}`);
                                        recognizer.close();

                                        // do some funky stuff with this text
                                        const textClient = new TextAnalyticsClient(process.env.AZ_COG_ENDPOINT as string, new AzureKeyCredential(process.env.AZ_COG_KEY as string));
                                        textClient.analyzeSentiment([result.text]).then(r => {
                                            _adapter.continueConversation(ref, async (ctx) => {
                                                await ctx.sendActivity(`The meeting had a ${(r[0] as any).sentiment} vibe`);
                                            });
                                        });
                                        const actions = {
                                            extractSummaryActions: [{ modelVersion: "latest", orderBy: "Rank", maxSentenceCount: 3 }]
                                        };
                                        const poller = await textClient.beginAnalyzeActions([result.text], actions, "en");
                                        poller.onProgress(() => {
                                            log(`Number of actions still in progress: ${poller.getOperationState().actionsInProgressCount}`);
                                        });

                                        const resultPages = await poller.pollUntilDone();

                                        // Show the summary of the meeting
                                        for await (const page of resultPages) {
                                            const extractSummaryAction = page.extractSummaryResults[0];
                                            if (!extractSummaryAction.error) {
                                                for (const doc of extractSummaryAction.results) {
                                                    if (!doc.error) {
                                                        let message = "**Meeting summary:**\n\n";
                                                        for (const sentence of doc.sentences) {
                                                            message += "- " + sentence.text + "\n";
                                                        }
                                                        _adapter.continueConversation(ref, async (ctx) => {
                                                            await ctx.sendActivity({ text: message, textFormat: "markdown" });
                                                        });
                                                    } else {
                                                        log("\tError:" + doc.error);
                                                    }
                                                }
                                            }
                                        }
                                    });
                                });

                                command.input(downloadPath);
                                command.audioChannels(2);
                                command.format("wav");
                                command.save(path.resolve(__dirname, "audio.wav"));
                            });
                            writer.on("error", () => {
                                log("Failed to write file");
                            });
                        });
                    });
                }
                if (video || transcript || texts) {
                    // start transcription of video
                    // const ctx = new media.AzureMediaServicesContext(creds, sub);
                    // const x = new media.Jobs(ctx);
                    // x.create(rg, acnt, transform, name, {
                    //     input:
                    // });
                    // media.Jobs
                    return {
                        statusCode: 200,
                        type: "application/vnd.microsoft.activity.message",
                        value: "Processing meeting information, brb..."
                    } as any;
                } else {
                    return {
                        statusCode: 200,
                        type: "application/vnd.microsoft.activity.message",
                        value: "Nothing to process, I hope you had a great meeting!"
                    } as any;
                }
            } catch (err) {
                log(`Error: ${err.message}`);
                return {
                    statusCode: 400,
                    type: "application/vnd.microsoft.error",
                    value: { message: "Processing failed!" }
                } as any;

            }

        }
        return {
            statusCode: 400,
            type: "application/vnd.microsoft.error",
            value: { message: "Invalid verb" }
        } as any;

    };

    public async sendWelcomeCard(context: TurnContext): Promise<void> {
        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        await context.sendActivity({ attachments: [welcomeCard] });
    }

    public async sendNewMeetingCard(context: TurnContext, data: any): Promise<void> {
        const card = CardFactory.adaptiveCard(NewMeetingCard(data));
        await context.sendActivity({ attachments: [card] });
    }

    public static async sendAfterMeetingCard(id: string, meetingId: string): Promise<void> {
        if (conversationReferences.exists(`/${id}`)) {
            const ref = conversationReferences.getData(`/${id}`) as Partial<ConversationReference>;
            _adapter.continueConversation(ref, async (context) => {
                const card = CardFactory.adaptiveCard(AfterMeetingCard(meetingId));
                await context.sendActivity({ attachments: [card] });
            });
        }
    }

    /**
     * Webhook for incoming calls
     */
    @BotCallingWebhook("/api/calling")
    public async onIncomingCall(req: express.Request, res: express.Response) {
        callingHandler(req, res);
    }

    public async sendOnlineMeetingStartMessage(chatInfo: any): Promise<void> {

    }

    public async handleTeamsSigninTokenExchange(context: TurnContext, query: SigninStateVerificationQuery): Promise<void> {
        log("handleTeamsSigninTokenExchange");
        if (await this._ssoOAuthHelper.shouldProcessTokenExchange(context)) {
            // nop
        } else {
            await this.dialog.run(context, this.dialogState);
        }
    }

    public async handleTeamsSigninVerifyState(context: TurnContext, query: SigninStateVerificationQuery): Promise<void> {
        await this.dialog.run(context, this.dialogState);
    }

}

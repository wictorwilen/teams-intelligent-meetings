import axios from "axios";
import debug = require("debug");
import express = require("express");
import jsonwebtoken = require("jsonwebtoken");
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta";
import meetings from "../meetingManager";
import chalk = require("chalk");
import { leaveCall } from "../api/commands";
import { MeetaiBot } from "./MeetaiBot";
const log = debug("msteams");

// tslint:disable-next-line: interface-over-type-literal
type JwtKeys = {
    keys: Array<{
        kty: string;
        use: string;
        kid: string;
        x5t: string;
        n: string;
        e: string;
        x5c: string[];
        endorsements: string[];
    }>;
};

/**
 * Fetches the keys from Azure AD
 * @returns the JWT keys
 */
const loadAzureADKeys = (): Promise<JwtKeys> => {
    return new Promise<any>((resolve, reject) => {
        axios.get("https://login.botframework.com/v1/.well-known/openidconfiguration").then(openidconfig => {
            axios.get(openidconfig.data.jwks_uri).then(result => {
                resolve(result.data);
            }).catch(err => {
                reject(err);
            });
        }).catch(err => {
            reject(err);
        });
    });
};

let botFxKeys: JwtKeys;
loadAzureADKeys().then(r => { botFxKeys = r; });

/**
 * VALIDATE INCOMING REQEUSTS
 * validataion as in https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/calls-and-meetings/call-notifications
 */
const validateRequest = (req: express.Request): number => {
    try {

        if (!req.headers.authorization) {
            return 401;
        }

        const token = req.headers.authorization!.split(" ")[1];
        const decoded = jsonwebtoken.decode(token, { complete: true }) as { [key: string]: any };
        const key = botFxKeys.keys.find(k => k.kid === decoded.header.kid);
        if (key) {
            if (jsonwebtoken.verify(
                token,
                `-----BEGIN CERTIFICATE-----\n${key.x5c[0]}\n-----END CERTIFICATE-----`,
                {
                    ignoreNotBefore: true // NOTE: sometimes we get a not before jwt not active error
                })) {
                // TODO: Move iss/aud to options above
                if (!(decoded.payload.iss === "https://api.botframework.com" &&
                    decoded.payload.aud === process.env.MICROSOFT_APP_ID)) {
                    // optionally add a verification for tid as well
                    log("Invalid iss or aud");
                    return 401;
                }
                log("Access token fully verified");
            } else {
                log("Could not verify JWT token");
                return 401;
            }
        } else {
            log("No key found to validate");
            return 401;
        }
    } catch (ex) {
        log(`Validating token failed with: ${ex}`);
        return 500;
    }
    return 0;
};

/**
 * Defines the calling webhook
 */
export default (req: express.Request, res: express.Response) => {
    log("Incoming call");

    // Validate the request
    const validation = validateRequest(req);
    if (validation !== 0) {
        res.status(validation).send();
        return;
    }

    const incoming: MicrosoftGraphBeta.CommsNotifications = req.body;
    let retval = 200;

    // Process the incoming request
    if (incoming && incoming.value) {
        incoming.value.forEach(async (incomingCall: MicrosoftGraphBeta.CommsNotification) => { // NOTE: there can be many call notifications
            log(incomingCall.changeType);

            if (incomingCall.resourceUrl) {
                // Extract details from the resourceUrl
                const meetingId = incomingCall.resourceUrl.split("/")[3]; // "resourceUrl": "/communications/calls/f31f5b00-b724-4c34-960c-3176607fd717/participants",
                const resourceData = (incomingCall as any).resourceData as MicrosoftGraphBeta.Call;
                const id = resourceData?.chatInfo?.threadId;

                // Check if we're managing this meeting
                let meeting = meetings.getById(meetingId);
                if (!meeting && id) {
                    meeting = meetings.getByThreadId(id);
                    if (meeting) {
                        meeting.id = meetingId;
                        meetings.update(meeting);
                    }
                }
                if (!meeting) {
                    log(chalk.cyan("Meeting not managed by this bot"));
                } else {

                    switch (incomingCall.changeType) {
                        case "deleted":
                            // meeting is ended
                            if (resourceData.state === "terminated") {
                                // Call was termindated
                                // inform the organizer
                                const organizer = (resourceData.meetingInfo as MicrosoftGraphBeta.OrganizerMeetingInfo).organizer;
                                if (organizer && organizer.user) {
                                    // organizer.user.id
                                    MeetaiBot.sendAfterMeetingCard(organizer.user.id as string, meetingId).then(() => {
                                        log(chalk.magenta("After meeting card sent"));
                                    }).catch(err => {
                                        log(`Unable to do after meeting stuff: ${err}`);
                                    });
                                }
                            };
                            break;
                        case "created":
                            break;
                        case "updated":

                            // These notifications are about ops on participants (mute, join etc)
                            if (incomingCall.resourceUrl?.endsWith("/participants")) {

                                // update participants in our local database
                                (resourceData as MicrosoftGraphBeta.Participant[]).forEach(p => {
                                    meetings.updateParticipant(meeting!, p);
                                });
                                meeting.activeParticipants = (resourceData as MicrosoftGraphBeta.Participant[]).length;

                                log(chalk.grey(`Active participants: ${meeting.activeParticipants}`));

                                // check if there is a TeamsRecorder
                                const recorder = (resourceData as MicrosoftGraphBeta.Participant[]).find(p => {
                                    return p.info && p.info.identity && p.info.identity.application && (p.info.identity.application as any).ApplicationType === "TeamsRecorder";
                                });
                                if (recorder) {
                                    log(chalk.blueBright("Recording is active!"));
                                    meeting.recording = true;
                                    // INFO: the actual meeting recording URL can be found in thread/messages via Graph. This requires access to protected apis if it has to be read using app permissions
                                };

                                meetings.update(meeting);

                                // only one participant - if it is the bot we should end the meeting
                                if (meeting.activeParticipants === 1) {
                                    const p = (resourceData as MicrosoftGraphBeta.Participant[])[0];
                                    if (p.info && p.info.identity && p.info.identity.application && p.info.identity.application.id === process.env.MICROSOFT_APP_ID) {
                                        log(chalk.grey("Only the bot in the meeting"));
                                        setTimeout(() => {
                                            const m = meetings.getById(meetingId);
                                            if (m?.activeParticipants === 1) {
                                                leaveCall((p.info!.identity!.application! as any).tenantId, meetingId);
                                            } else {
                                                log("Aborting auto-leave");
                                            }
                                        }, 60 * 1000); // 1 minute
                                    }
                                }

                                retval = 200;
                                // Check if anyone has a published state (raise hands for now)
                                const publishedStates = (resourceData as MicrosoftGraphBeta.Participant[]).filter(p => {
                                    return (p as any).publishedStates && (p as any).publishedStates.length !== 0;
                                });

                                // Check if anyone has their hands raised
                                publishedStates.forEach(p => {
                                    const publishedState: { type: "raiseHand" }[] = (p as any).publishedStates;
                                    if (publishedState.filter(x => x.type === "raiseHand").length > 0) {
                                        log(`${p.info?.identity?.user?.displayName} raised the hand`);
                                    }
                                });

                                // TODO: check for mute/unmute

                            } else if (incomingCall.resourceUrl?.endsWith("/operations")) {
                                // ignore here
                                retval = 200;
                            } else {
                                // Answer the call
                                switch (resourceData.state) {
                                    case "establishing": // [1] This is the first incoming ops on an incoming call
                                        retval = 200;
                                        break;
                                    case "established": // [2] This is the second part happening for incoming calls (must answer within 15 secs)
                                        retval = 200;
                                        // Do after join meeting stuff
                                        break;
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        });
        res.status(retval).send();
    } else {
        res.sendStatus(401);
    }
};

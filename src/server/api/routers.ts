import { way } from "expressways";
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider, TokenCredentialAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import express = require("express");
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta";
import "isomorphic-fetch";
import debug = require("debug");
import chalk = require("chalk");
import { tenantsDb } from "./public";
import Meetings from "../meetingManager";

import * as Incidents from "../incidents.json";

const log = debug("msteams");

const random = (length = 8) => {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    let str = "";
    for (let i = 0; i < length; i++) {
        str += chars.charAt(Math.floor(Math.random() * chars.length));
    }

    return str;
};

class XTokenCredentialAuthenticationProvider extends TokenCredentialAuthenticationProvider {
    public async getAccessToken(): Promise<string> {
        const token = await super.getAccessToken();
        log(chalk.grey(`token: ${token}`));
        return token;
    }
}

export const getGraphClient = (tid: string, scopes: string[]) => {
    const credential = new ClientSecretCredential(tid, process.env.MICROSOFT_APP_ID as string, process.env.MICROSOFT_APP_PASSWORD as string);
    const authProvider = new XTokenCredentialAuthenticationProvider(credential, { scopes });

    const client = Client.initWithMiddleware({
        debugLogging: true,
        authProvider
    });
    return client;
};

/**
 * Creates a new Online meeting.
 * Users have to manually join using the Join URL
 */
export const get1 = way("put", "/meetings", async (req: express.Request, res, next) => {
    const d = new Date();
    const d2 = new Date();
    d2.setHours(d.getHours() + 1);
    const incident = Incidents[Math.floor(Math.random() * Incidents.length)];
    const body: MicrosoftGraphBeta.OnlineMeeting = {
        startDateTime: d.toISOString(),
        endDateTime: d2.toISOString(),
        subject: `❗Incident: '${incident.title} for ${incident.company}`,
        externalId: "incident-" + incident.id + "-" + random(4), // use this to reuse meetings
        recordAutomatically: true,
        isEntryExitAnnounced: true,
        participants: {
            organizer: {
                identity: {
                    application: {
                        id: process.env.MICROSOFT_APP_ID,
                        displayName: "meet.ai"
                    }
                }
            },
            attendees: [
                {
                    identity: {
                        user: {
                            id: (req.user as any).id
                        }
                    }
                }
            ]
        }
    };

    // Create the online meeting with an external id
    const client = getGraphClient((req.user as any)?.tid, ["https://graph.microsoft.com/.default"]);
    client.api(`/users/${(req.user as any).id}/onlineMeetings/createOrGet`).version("beta").post(body).then(async (result: MicrosoftGraphBeta.OnlineMeeting) => {
        if (!result) {
            throw new Error("No result");
        }

        // return the join URL
        res.status(201).header("location", result.joinUrl!).send("Created");

        // register the meeting
        Meetings.add({
            id: undefined,
            threadId: result.chatInfo?.threadId as string,
            data: {
                incident: incident.id
            },
            participants: [],
            activeParticipants: 0
        });

        // get the "real" app id
        const tid = (req.user as any)?.tid;
        const teamsAppId = tenantsDb.getData(`/${tid}/teamsAppId`);

        // Install the app to the meeting
        log(`App id: ${teamsAppId}`);
        client.api(`/chats/${result.chatInfo!.threadId}/installedApps`)
            .version("beta")
            .post({
                "teamsApp@odata.bind": `https://graph.microsoft.com/beta/appCatalogs/teamsApps/${teamsAppId}`
            })
            .then(res => {
                log("Success");
                // Join the bot to the call
                const call = {
                    "@odata.type": "#microsoft.graph.call",
                    callbackUri: `https://${process.env.PUBLIC_HOSTNAME}/api/calling`,
                    requestedModalities: [
                        "audio"
                    ],
                    mediaConfig: {
                        "@odata.type": "#microsoft.graph.serviceHostedMediaConfig",
                        preFetchMedia: [
                            {
                                uri: `https://${process.env.PUBLIC_HOSTNAME}/assets/beep.wav`,
                                resourceId: "f8971b04-b53e-418c-9222-c82ce681a582"
                            },
                            {
                                uri: `https://${process.env.PUBLIC_HOSTNAME}/assets/cool.wav`,
                                resourceId: "86dc814b-c172-4428-9112-60f8ecae1edb"
                            }
                        ]
                    },
                    chatInfo: {
                        "@odata.type": "#microsoft.graph.chatInfo",
                        threadId: result.chatInfo?.threadId,
                        messageId: "0"
                    },
                    meetingInfo: {
                        "@odata.type": "#microsoft.graph.organizerMeetingInfo",
                        organizer: {
                            "@odata.type": "#microsoft.graph.identitySet",
                            user: {
                                "@odata.type": "#microsoft.graph.identity",
                                id: (req.user as any).id,
                                tenantId: (req.user as any)?.tid, // REQUIRED,
                                displayName: "Wictor"
                            }
                        },
                        allowConversationWithoutHost: true
                    },
                    tenantId: (req.user as any)?.tid // REQUIRED
                };
                client
                    .api("/communications/calls")
                    .version("beta")
                    .post(call)
                    .then(callResult => {
                        log("Calling bot worked as expected");

                        // Add the tab
                        client.api(`/chats/${result.chatInfo!.threadId}/tabs`)
                            .version("beta")
                            .post({
                                displayName: "Incident report",
                                "teamsApp@odata.bind": `https://graph.microsoft.com/beta/appCatalogs/teamsApps/${teamsAppId}`,
                                configuration: {
                                    entityId: incident.id,
                                    contentUrl: `https://${process.env.PUBLIC_HOSTNAME}/meetAiTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}&incident=${incident.id}`,
                                    websiteUrl: `https://${process.env.PUBLIC_HOSTNAME}/meetAiTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}&incident=${incident.id}`,
                                    removeUrl: `https://${process.env.PUBLIC_HOSTNAME}/meetAiTab/remove.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
                                }
                            } as MicrosoftGraphBeta.TeamsTab)
                            .then(addTabResult => {
                                log(chalk.green("Tab added successfully"));
                            })
                            .catch(err => {
                                log(chalk.red(`Adding tab did not work:${err.message}`));
                            });
                    })
                    .catch(err => {
                        log(chalk.red(`Calling bot did not work:${err.message}`));
                    });
            })
            .catch(err => {
                log(chalk.red(`Failed adding app: ${err.message}`));
            });

    }).catch(err => {
        if (err.message === "Application does not have permission to CreateOrGet online meeting on behalf of this user.") {
            // See: https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy
            /*
            * Install-Module -Name MicrosoftTeams -Force -AllowClobber -Scope CurrentUser
            * Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
            * Import-Module MicrosoftTeams
            * Connect-MicrosoftTeams
            * New-CsApplicationAccessPolicy -Identity Meet-ai-meeting-policy -AppIds "b398f076-4b0f-43c6-accc-e180d8262f5a" -Description "Meet.ai schedule meetings"
            * Grant-CsApplicationAccessPolicy -PolicyName Meet-ai-meeting-policy -Identity "f4b157b0-3e3e-410c-9648-b9bd5d53a689"
            * (all tenant) Grant-CsApplicationAccessPolicy -PolicyName Test-policy -Global
            */
            res.status(401).send(err.message);
        } else {
            res.status(500).send(err.message);
            log(chalk.red(err.message));
        }
    });
});

export const call = way("put", "/calls", (req: express.Request, res, next) => {

    const body: MicrosoftGraphBeta.Call = {
        callbackUri: `https://${process.env.PUBLIC_HOSTNAME}/api/calling`,
        requestedModalities: [
            "audio",
            "video"
        ],
        meetingInfo: {
            allowConversationWithoutHost: true
        },
        mediaConfig: {
            removeFromDefaultAudioGroup: true
        },
        tenantId: (req.user as any)?.tid // REQUIRED
    };

    // Create the online meeting with an external id
    const client = getGraphClient((req.user as any)?.tid, ["https://graph.microsoft.com/.default"]);
    client.api("/communications/calls").version("beta").post({
        "@odata.type": "#microsoft.graph.call",
        callbackUri: `https://${process.env.PUBLIC_HOSTNAME}/api/calling`,
        targets: [
            {
                "@odata.type": "#microsoft.graph.invitationParticipantInfo",
                identity: {
                    "@odata.type": "#microsoft.graph.identitySet",
                    user: {
                        "@odata.type": "#microsoft.graph.identity",
                        id: (req.user as any).id
                    }
                }
            },
            {
                "@odata.type": "#microsoft.graph.invitationParticipantInfo",
                identity: {
                    "@odata.type": "#microsoft.graph.identitySet",
                    user: {
                        "@odata.type": "#microsoft.graph.identity",
                        id: "8edf4fce-cf66-4429-9f8b-544e4ef9cd00"
                    }
                }
            }
        ],
        subject: "Important meeting",
        requestedModalities: [
            "audio"
        ],
        mediaConfig: {
            "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
        },
        tenantId: (req.user as any)?.tid // REQUIRED
    }).then((result) => {
        log(result);
        res.status(201).send("Created");
    }).catch(err => {
        if (err.message === "Application does not have permission to CreateOrGet online meeting on behalf of this user.") {
            // See: https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy
            /*
            * Install-Module -Name MicrosoftTeams -Force -AllowClobber -Scope CurrentUser
            * Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
            * Import-Module MicrosoftTeams
            * Connect-MicrosoftTeams
            * New-CsApplicationAccessPolicy -Identity Meet-ai-meeting-policy -AppIds "b398f076-4b0f-43c6-accc-e180d8262f5a" -Description "Meet.ai schedule meetings"
            * Grant-CsApplicationAccessPolicy -PolicyName Meet-ai-meeting-policy -Identity "f4b157b0-3e3e-410c-9648-b9bd5d53a689"
            * (all tenant) Grant-CsApplicationAccessPolicy -PolicyName Test-policy -Global
            */
            res.status(401).send(err.message);
        } else {
            res.status(500).send(err.message);
            log(chalk.red(err.message));
        }
    });
});

export const teamsAppId = way<any, { id: string }>("post", "/teamsapp/id", (req: express.Request, res, next) => {
    const tid = (req.user as any)?.tid;
    const id = req.body.id;

    if (tenantsDb.exists(`/${tid}`)) {
        const data = tenantsDb.getData(`/${tid}`);
        data.teamsAppId = id;
        tenantsDb.push(`/${tid}`, data);
        res.status(209).send();
    } else {
        res.status(401).send();
    }
});

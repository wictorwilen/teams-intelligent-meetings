import * as React from "react";
import { Provider, Flex, Header, Input, Dropdown } from "@fluentui/react-northstar";
import { useState, useEffect, useRef, useCallback } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, pages, authentication } from "@microsoft/teams-js";
import * as teamsFx from "@microsoft/teamsfx";
import jwtDecode from "jwt-decode";
import axios from "axios";
/**
 * Implementation of meet.ai configuration page
 */
export const MeetAiTabConfig = () => {

    const [{ inTeams, theme, context }] = useTeams({});
    const [id, setId] = useState<string>();
    const [incidents, setIncidents] = useState<any[]>([]);

    const onSaveHandler = useCallback((saveEvent: pages.config.SaveEvent) => {
        const host = "https://" + window.location.host;
        pages.config.setConfig({
            contentUrl: host + "/meetAiTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}&incident=" + id,
            websiteUrl: host + "/meetAiTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}&incident=" + id,
            suggestedDisplayName: "meet.ai",
            removeUrl: host + "/meetAiTab/remove.html?theme={theme}",
            entityId: id
        }).then(() => {
            saveEvent.notifySuccess();
        }).catch(err => {
            saveEvent.notifyFailure(err);
        });
    }, [id]);

    useEffect(() => {
        if (inTeams) {
            pages.config.registerOnSaveHandler(onSaveHandler);
        }
    }, [inTeams, onSaveHandler]);

    useEffect(() => {
        if (inTeams) {
            pages.config.setValidityState(false);

            authentication.getAuthToken({
                silent: true,
                resources: [process.env.TAB_APP_URI as string]
            }).then(token => {
                const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                teamsFx.loadConfiguration({
                    authentication: {
                        initiateLoginEndpoint: `https://${process.env.PUBLIC_HOSTNAME}/ile`,
                        clientId: process.env.TAB_APP_ID,
                        tenantId: (decoded as any).tid,
                        authorityHost: "https://login.microsoftonline.com",
                        applicationIdUri: process.env.TAB_APP_URI,
                        simpleAuthEndpoint: `https://${process.env.PUBLIC_HOSTNAME}`
                    }
                });

                axios.get("/_api/incidents", { headers: { Authorization: `Bearer ${token}` } }).then(result => {
                    if (result.status === 200) {
                        setIncidents(result.data);
                    }
                });
                app.notifySuccess();
            }).catch(message => {

                app.notifyFailure({
                    reason: app.FailedReason.AuthFailed,
                    message
                });
            });
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setId(context.page.id);
            pages.config.setValidityState(true);
            app.notifySuccess();
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [context]);

    useEffect(() => {
        if (id) {
            pages.config.setValidityState(true);
        } else {
            pages.config.setValidityState(false);
        }
    }, [id]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <Header content="Select incident" />
                        <Dropdown
                            items={incidents.map(i => ({
                                header: i.id + " : " + i.title,
                                selected: i.id === id,
                                id: i.id
                            }))}
                            onChange={(e, data) => {
                                if (data && data.value) {
                                    setId((data.value as any).id);
                                } else {
                                    setId(undefined);
                                }
                            }}
                        />
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};

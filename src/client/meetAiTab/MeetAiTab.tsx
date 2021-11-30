import * as React from "react";
import { Provider, Flex, Text, Button, Header, Card, Image } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams, getQueryVariable } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import axios from "axios";
import * as teamsFx from "@microsoft/teamsfx";

/**
 * Implementation of the meet.ai content page
 */
export const MeetAiTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [token, setToken] = useState<string>();
    const [tid, setTid] = useState<string>();
    const [error, setError] = useState<string>();
    const [joinUrl, setJoinUrl] = useState<string | undefined>(undefined);
    const [consentLoading, setConsentLoading] = useState(false);
    const [incident, setIncident] = useState<any>({});

    useEffect(() => {
        if (inTeams) {
            authentication.getAuthToken({
                silent: true,
                resources: [process.env.TAB_APP_URI as string]
            }).then(token => {
                const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                setName(decoded!.name);
                setToken(token);
                setTid(decoded!.tid);
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
                app.notifySuccess();
            }).catch(message => {
                setError(message);
                app.notifyFailure({
                    reason: app.FailedReason.AuthFailed,
                    message
                });
            });
        }
    }, [inTeams]);

    useEffect(() => {
        if (context && token) {
            setEntityId(context.page.id);
            if (context.page.id === undefined) {
                // fallback to url value
                setEntityId(getQueryVariable("incident"));
            }
            if (entityId !== undefined) {
                axios.get("/_api/incidents/" + entityId, { headers: { Authorization: `Bearer ${token}` } }).then(result => {
                    if (result.status === 200) {
                        setIncident(result.data);
                    } else {
                        setError("Incident not found");
                    }
                });
            }
        };
    }, [context, entityId, token]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="Incident manager" />
                </Flex.Item>
                <Flex.Item>
                    {incident &&
                        <Card>
                            <Card.Header>
                                <Flex column>
                                    <Text content={incident.title} weight="bold" />
                                    <Text content={incident.company} size="small" />
                                </Flex>
                            </Card.Header>
                            <Card.Body>
                                <Flex column gap="gap.small">
                                    <Image src={incident.picture} fluid />
                                    <Text content={incident.description} />
                                </Flex>
                            </Card.Body>
                        </Card>
                    }
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Wictor Wilén" />
                </Flex.Item>
            </Flex>
        </Provider >
    );
};

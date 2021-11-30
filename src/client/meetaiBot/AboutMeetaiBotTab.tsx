import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import axios from "axios";
import * as teamsFx from "@microsoft/teamsfx";

/**
 * Implementation of the aboutMeetaiBot content page
 */
export const AboutMeetaiBotTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [name, setName] = useState<string>();
    const [token, setToken] = useState<string>();
    const [tid, setTid] = useState<string>();
    const [error, setError] = useState<string>();
    const [joinUrl, setJoinUrl] = useState<string | undefined>(undefined);
    const [consentLoading, setConsentLoading] = useState(false);

    useEffect(() => {
        if (inTeams === true) {
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

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="Meeting demo setup and config" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={`Hello ${name}`} />
                        </div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                            <Header as="h2">Online meeting</Header>
                            <Button onClick={() => {
                                setJoinUrl(undefined);
                                axios.put("/_api/meetings", {}, { headers: { Authorization: `Bearer ${token}` } }).then(result => {
                                    if (result.status === 201) {
                                        setJoinUrl(result.headers.location);
                                    }
                                });
                            }}>Create an online meeting</Button>
                            {joinUrl && <a href={joinUrl} target="_new">Click here to join meeting</a>}
                        </div>

                        <div>
                            <Button onClick={() => {
                                axios.put("/_api/calls", {}, { headers: { Authorization: `Bearer ${token}` } });
                            }}>Create a call</Button>
                        </div>
                        <div>
                            <Button style={{ backgroundColor: "darkRed", color: "white" }} loading={consentLoading} onClick={async () => {
                                setConsentLoading(true);
                                authentication.authenticate({
                                    url: `https://login.microsoftonline.com/${tid}/adminconsent?client_id=${process.env.MICROSOFT_APP_ID}&redirect_uri=https://${process.env.PUBLIC_HOSTNAME}/_api/consent/bot`
                                }).then(async () => {
                                    const credential = new teamsFx.TeamsUserCredential();
                                    const graphClient = teamsFx.createMicrosoftGraphClient(credential, ["AppCatalog.Read.All"]);
                                    const response = await graphClient.api(`/appCatalogs/teamsApps?$filter=externalId eq '${process.env.APPLICATION_ID}'`).get();
                                    setConsentLoading(false);
                                }).catch(err => {
                                    alert(err);
                                    setConsentLoading(false);
                                });
                            }}>Admin: consent bot app</Button>
                            <Button style={{ backgroundColor: "darkRed", color: "white" }} loading={consentLoading} onClick={async () => {
                                setConsentLoading(true);
                                authentication.authenticate({
                                    url: `https://login.microsoftonline.com/${tid}/adminconsent?client_id=${process.env.TAB_APP_ID}&redirect_uri=https://${process.env.PUBLIC_HOSTNAME}/_api/consent/tab`
                                }).then(async (result) => {
                                    const credential = new teamsFx.TeamsUserCredential();
                                    const graphClient = teamsFx.createMicrosoftGraphClient(credential, ["AppCatalog.Read.All"]);
                                    const response = await graphClient.api(`/appCatalogs/teamsApps?$filter=externalId eq '${process.env.APPLICATION_ID}'`).get();
                                    console.log(response);
                                    if (response.value.length !== 1) {
                                        alert("App not found");
                                        setConsentLoading(false);
                                    } else {
                                        const id = response.value[0].id;
                                        axios.post("/_api/teamsapp/id", { id }, { headers: { Authorization: `Bearer ${token}` } }).then(result => {
                                            if (response.status === 201) {
                                                // nop
                                            }
                                            setConsentLoading(false);
                                            alert("All set and done");
                                        });
                                    }
                                }).catch(err => {
                                    alert(err);
                                    setConsentLoading(false);
                                });
                            }}>Admin: consent and setup tab app</Button>
                        </div>
                    </div>
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

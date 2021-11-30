import express = require("express");
import { BearerStrategy, IBearerStrategyOptionWithRequest, ITokenPayload, VerifyCallback } from "passport-azure-ad";
export default () => new BearerStrategy(
    {
        identityMetadata: "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
        clientID: process.env.TAB_APP_ID as string,
        audience: [
            process.env.TAB_APP_ID as string,
            process.env.TAB_APP_URI as string
        ],
        loggingLevel: "error",
        validateIssuer: false,
        passReqToCallback: true
    } as IBearerStrategyOptionWithRequest,
    async (request: express.Request, token: ITokenPayload, done: VerifyCallback) => {
        done(null, { tid: token.tid, name: token.name, scp: token.scp, id: token.oid }, token);
    });

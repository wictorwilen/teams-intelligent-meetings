import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";
import * as compression from "compression";
import * as appInsights from "applicationinsights";
import { SimpleAuthRouter } from "teams-simpleauth";
import api from "./api";
// Initialize debug logging module
const log = debug("msteams");

log("Initializing Microsoft Teams Express hosted App...");

// Initialize dotenv, to use .env file settings if existing
require("dotenv").config();

// Set up app insights
appInsights.setup(process.env.APPINSIGHTS_INSTRUMENTATIONKEY).start();
// The import of components has to be done AFTER the dotenv config
// eslint-disable-next-line import/first
import * as allComponents from "./TeamsAppsComponents";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
express.use(Express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
        (req as any).rawBody = buf.toString();
    }
}));
express.use(Express.urlencoded({ extended: true }));

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Add compression - uncomment to remove compression
express.use(compression());

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));

express.use("/auth/token", SimpleAuthRouter({
    appId: process.env.TAB_APP_ID as string,
    appIdUri: process.env.TAB_APP_URI as string,
    appSecret: process.env.TAB_APP_SECRET as string,
    scopes: ["AppCatalog.Read.All"],
    tenantId: "common"
}));

// add API's
express.use("/_api", api());

// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));

// Set the port
express.set("port", port);

// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});
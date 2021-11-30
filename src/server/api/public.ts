/* eslint-disable camelcase */
import { way } from "expressways";
import express = require("express");
import debug = require("debug");

import chalk = require("chalk");
import { JsonDB } from "node-json-db";
import { Config } from "node-json-db/dist/lib/JsonDBConfig";

const log = debug("msteams");

// tenantsDb contains all the tenants using the app
export const tenantsDb = new JsonDB(new Config("tenants.db.json", true, false, "/"));

/**
 * API endpoint used for redirect uri for the bot
 */
export const consentBot = way<{ id: string }, any, any, { admin_consent: string, tenant: string }>("get", "/consent/bot", (req, res, next) => {
    log(req.url);
    if (req.query.admin_consent === "True") {
        const tenantId = req.query.tenant;
        const tenant = tenantsDb.exists(`/${tenantId}`);
        if (tenant) {
            log(chalk.green("Tenant already registered"));
            const data = tenantsDb.getData(`/${tenantId}`);
            data.bot = true;
            tenantsDb.push(`/${tenantId}`, data);
        } else {
            log(chalk.yellow("New tenant registered"));
            tenantsDb.push(`/${tenantId}`, { bot: true });
        }
        res.redirect("/meetAiTab/authEnd.html");
    } else {
        res.status(500).send("Invalid command");
    }
});

/**
 * API endpoint used for redirect uri for the tab
 */
export const consentTab = way<{ id: string }, any, any, { admin_consent: string, tenant: string }>("get", "/consent/tab", (req, res, next) => {
    log(req.url);
    if (req.query.admin_consent === "True") {
        const tenantId = req.query.tenant;
        const tenant = tenantsDb.exists(`/${tenantId}`);
        if (tenant) {
            log(chalk.green("Tenant already registered"));
            const data = tenantsDb.getData(`/${tenantId}`);
            data.tab = true;
            tenantsDb.push(`/${tenantId}`, data);
        } else {
            log(chalk.yellow("New tenant registered"));
            tenantsDb.push(`/${tenantId}`, { tab: true });
        }
        res.redirect("/meetAiTab/authEnd.html");
    } else {
        res.status(500).send("Invalid command");
    }
});

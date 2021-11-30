/* eslint-disable camelcase */
import { way } from "expressways";
import express = require("express");
import debug = require("debug");

import * as incidents from "../incidents.json";

const log = debug("msteams");

export const consentBot = way<never, never, any>("get", "/incidents", (req, res, next) => {
    res.send(incidents);
});

export const consentTab = way<{ id: string }, never, any>("get", "/incidents/:id", (req, res, next) => {
    const incident = incidents.filter(i => i.id === req.params.id);
    if (incident && incident.length === 1) {
        res.send(incident[0]);
    } else {
        res.status(404).send();
    }
});

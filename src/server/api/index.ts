import debug = require("debug");
import express = require("express");
import { expressways } from "expressways";
import passport = require("passport");
import clientBearerStrategy from "../bearerStrategy";
import * as routes from "./routers";
import * as incidents from "./incidents";
import * as publicRoutes from "./public";
const log = debug("msteams");

/**
 * Sets up default APIs
 */
export default (): express.Router => {
    const router = express.Router();

    const pass = new passport.Passport();
    router.use(pass.initialize());
    pass.use(clientBearerStrategy());

    router.use((req, res, next) => {
        res.header("Cache-control", "no-store, no-cache, must-revalidate, private");
        res.header("Access-Control-Allow-Origin", "*");
        res.header("Access-Control-Allow-Headers", "Authorization, Origin, X-Requested-With, Content-Type, Accept");
        next();
    });

    // secured API's
    expressways({
        router,
        ways: incidents,
        log,
        handlers: [
            pass.authenticate("oauth-bearer", { session: false })
        ]
    });
    expressways({
        router,
        ways: routes,
        log,
        handlers: [
            pass.authenticate("oauth-bearer", { session: false })
        ]
    });

    // public API's
    expressways({
        router,
        ways: publicRoutes,
        log
    });

    return router;
};

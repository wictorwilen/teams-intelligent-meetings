import { app, authentication } from "@microsoft/teams-js";

/**
 * Implementation of the meet.ai auth end page
 */
export const AuthEnd = () => {
    app.initialize().then(() => {
        authentication.notifySuccess("true");
    });
};

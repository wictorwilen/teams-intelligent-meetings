import chalk = require("chalk");
import { log } from "debug";
import { getGraphClient } from "./routers";

export const leaveCall = (tid: string, callId: string): Promise<void> => {
    return new Promise((resolve, reject) => {
        const client = getGraphClient(tid, ["https://graph.microsoft.com/.default"]);
        client.api(`/communications/calls/${callId}`).version("beta").delete().then(() => {
            log(chalk.red("Left call"));
            resolve();
        }).catch(err => {
            log(chalk.red(`Unable to leave call: ${err.message}`));
            reject(err);
        });
    });

};

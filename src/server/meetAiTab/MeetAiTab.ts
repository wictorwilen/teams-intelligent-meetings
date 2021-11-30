import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/meetAiTab/index.html")
@PreventIframe("/meetAiTab/config.html")
@PreventIframe("/meetAiTab/remove.html")
export class MeetAiTab {
}

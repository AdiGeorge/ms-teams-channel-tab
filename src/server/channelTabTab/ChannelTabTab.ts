import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/channelTabTab/index.html")
@PreventIframe("/channelTabTab/config.html")
@PreventIframe("/channelTabTab/remove.html")
export class ChannelTabTab {
}

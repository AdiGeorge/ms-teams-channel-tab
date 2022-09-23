import * as React from "react";
import { Provider, Flex, Header, Input, DropdownProps, Dropdown} from "@fluentui/react-northstar";
import { useState, useEffect, useRef } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, pages } from "@microsoft/teams-js";
import * as microsoftTeams from "@microsoft/teams-js";


/**
 * Implementation of channel tab Tab configuration page
 */
export const ChannelTabTabConfig = () => {

    const [{ inTeams, theme, context }] = useTeams({});
    const [text, setText] = useState<string>();
    const entityId = useRef("");
    const [mathOperator, setMathOperator] = useState<string>();


    const onSaveHandler = (saveEvent: pages.config.SaveEvent) => {
        const host = "https://" + window.location.host;
        pages.config.setConfig({
            contentUrl: host + "/channelTabTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
            websiteUrl: host + "/channelTabTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
            suggestedDisplayName: "channel tab Tab",
            removeUrl: host + "/channelTabTab/remove.html?theme={theme}",
            entityId: entityId.current
        }).then(() => {
            saveEvent.notifySuccess();
        });
    };

    useEffect(() => {
        if (context) {
            setMathOperator(context.entityId.replace("MathPage", ""));
            entityId.current = context.entityId;
            microsoftTeams.settings.registerOnSaveHandler(onSaveHandler);
            microsoftTeams.settings.setValidityState(true);
            microsoftTeams.appInitialization.notifySuccess();
            }
            // eslint-disable-next-line react-hooks/exhaustive-deps
          }, [context]);
    return (
        <Provider theme={theme}>
        <Flex gap="gap.smaller" style={{ height: "300px" }}>
          <Dropdown placeholder="Select the math operator"
            items={[
              "add",
              "subtract",
              "multiply",
              "divide"
            ]}
            onChange={(e, data) => {
              if (data) {
                const op = (data.value) ? data.value.toString() : "add";
                setMathOperator(op);
                entityId.current = `${op}MathPage`;
              }
            }}
            value={mathOperator}></Dropdown>
        </Flex>
      </Provider>
    );
};

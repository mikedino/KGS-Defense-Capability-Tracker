import * as React from "react";
import { Icon, IconButton, Stack, Text } from "@fluentui/react";
import { useHistory } from "react-router-dom";
import Strings from "../common/strings";
import styles from "../Dct.module.scss";
import { Security } from "../services/Security";
import { routes } from "../routing/routes";

const webViewPreferenceKey = "KGSDCTracker.WebView";

export const AppHeader: React.FC = () => {
    const history = useHistory();

    const isWebView = React.useMemo(() => {
        if (typeof window === "undefined") return false;

        const url = new URL(window.location.href);
        return url.searchParams.get("env")?.toLowerCase() === "webview";
    }, []);

    const toggleWebView = React.useCallback(() => {
        const url = new URL(window.location.href);

        if (isWebView) {
            url.searchParams.delete("env");
            window.sessionStorage.removeItem(webViewPreferenceKey);
        } else {
            url.searchParams.set("env", "WebView");
            window.sessionStorage.setItem(webViewPreferenceKey, "true");
        }

        window.location.assign(url.toString());
    }, [isWebView]);

    return (
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.mainHeader}>
            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                <Icon iconName="WebAppBuilderFragment" className={styles.titleIcon} />
                <Text
                    variant="xxLarge"
                    className={styles.title}
                    onClick={() => history.push(routes.home)}
                >
                    {Strings.ProjectName}
                </Text>
            </Stack>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                <IconButton
                    className={styles.headerIconButton}
                    iconProps={{ iconName: isWebView ? "BackToWindow" : "OpenInNewWindow" }}
                    title={isWebView ? "Exit WebView" : "Open in WebView"}
                    ariaLabel={isWebView ? "Exit WebView" : "Open in WebView"}
                    onClick={toggleWebView}
                />
                <Text className={styles.headerRole}>Role: {Security.RoleDisplay}</Text>
            </Stack>
        </Stack>
    );
};

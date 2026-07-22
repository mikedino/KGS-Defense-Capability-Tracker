import * as React from "react"
import { useState } from "react"
import { Stack, Text, Pivot, PivotItem, CommandBarButton } from "@fluentui/react"
import { customPivotStyles } from "../ui/ComponentStyles"
import { WebPartContext } from "@microsoft/sp-webpart-base"
import { ConfigManagement } from "./Configuration"
import { PersonnelManagement } from "./PersonnelManagement"
import styles from "../Dct.module.scss"
import { AppHeader } from "../ui/AppHeader"

interface IAdminPanelProps {
    context: WebPartContext
    onBack: () => void
}

export const AdminPanel: React.FunctionComponent<IAdminPanelProps> = ({ context, onBack }) => {
    const [selectedTab, setSelectedTab] = useState<string>("personnel");

    return (
        <Stack>
            <AppHeader />

            <Stack tokens={{ childrenGap: 20 }} className={styles.pageContent}>
                <Stack horizontalAlign="start">
                    <CommandBarButton text="Back (Home)" iconProps={{ iconName: "NavigateBack" }} onClick={onBack} />
                </Stack>
                <Stack tokens={{ childrenGap: 14 }}>
                    <Text variant="xxLarge">Administration</Text>
                    <Text variant="medium" style={{ color: "#666" }}>
                        Manage personnel and configuration lookup data
                    </Text>
                </Stack>

                <div className={styles.navSurface}>
                    <Pivot
                        styles={customPivotStyles}
                        selectedKey={selectedTab}
                        onLinkClick={(item) => setSelectedTab(item?.props.itemKey || "personnel")}
                        linkFormat="tabs"
                    >
                        <PivotItem headerText="Personnel" title="Manage personnel" ariaLabel="Manage personnel" itemKey="personnel" />
                        <PivotItem headerText="Configuration" title="Manage configuration" ariaLabel="Manage configuration" itemKey="configuration" />
                    </Pivot>
                </div>

                {selectedTab === "personnel" && <PersonnelManagement context={context} />}
                {selectedTab === "configuration" && <ConfigManagement />}
            </Stack>

        </Stack>
    )
}

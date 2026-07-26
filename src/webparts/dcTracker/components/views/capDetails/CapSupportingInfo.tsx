import * as React from "react";
import { Label, Stack, Text } from "@fluentui/react";
import { ICapabilityItem } from "../../common/props";
import styles from "../../Dct.module.scss";

interface ISupportingInfoProps {
    capState: ICapabilityItem;
}

const FieldDisplay: React.FC<{ label: string; value?: string; multiline?: boolean }> = ({
    label,
    value,
    multiline
}) => (
    <Stack style={{ minWidth: multiline ? "100%" : 220, flexGrow: multiline ? 1 : 0 }}>
        <Label>{label}</Label>
        <Text styles={{ root: { whiteSpace: multiline ? "pre-wrap" : "normal" } }}>
            {value || "—"}
        </Text>
    </Stack>
);

export const SupportingInfo: React.FC<ISupportingInfoProps> = ({ capState }) => {
    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} className={styles.detailCard}>
                <FieldDisplay label="Solution Type" value={capState.solutionType} />
                <FieldDisplay label="Platform" value={capState.platform} />
                <FieldDisplay label="Hosting Environment" value={capState.hostingEnv} />
                <FieldDisplay label="Connectivity" value={capState.connectivity} />
                <FieldDisplay label="Compliance" value={capState.compliance} />
                <FieldDisplay label="License Required?" value={capState.licenseReqd} />
                <FieldDisplay label="Coding Language" value={capState.codeLanguage} />
                <FieldDisplay label="Backend" value={capState.backend} />
                <FieldDisplay label="Server Requirements" value={capState.serverReqmts} />
                <FieldDisplay label="Licensing Requirements" value={capState.licenseReqmts} multiline />
                <FieldDisplay label="APIs/Extensibility" value={capState.extensibility} multiline />
            </Stack>
        </Stack>
    );
};

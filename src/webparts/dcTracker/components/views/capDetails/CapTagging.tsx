import * as React from "react";
import { Label, Link, Stack, Text } from "@fluentui/react";
import { ICapabilityItem, IOpportunityItem, IPastPerformanceItem, IProposalItem } from "../../common/props";
import styles from "../../Dct.module.scss";

interface ITaggingInfoProps {
    capState: ICapabilityItem;
}

const emptyTextStyles = { root: { color: "gray", fontStyle: "italic" } };

const getCapabilityAreas = (tag: IPastPerformanceItem): string =>
    tag.Capability_x0020_Area?.results?.length ? tag.Capability_x0020_Area.results.join(", ") : "—";

const FieldDisplay: React.FC<{ label: string; value?: string | number; className?: string }> = ({ label, value, className }) => (
    <Stack className={`${styles.tagDetailField} ${className ?? ""}`}>
        <Label>{label}</Label>
        <Text title={value ? String(value) : undefined}>{value || "—"}</Text>
    </Stack>
);

const TagLink: React.FC<{ href?: string; children: React.ReactNode }> = ({ href, children }) =>
    href ? (
        <Link href={href} target="_blank" rel="noopener noreferrer" title={String(children)}>
            {children}
        </Link>
    ) : (
        <Text>{children}</Text>
    );

const LinkFieldDisplay: React.FC<{
    label: string;
    href?: string;
    value: React.ReactNode;
    className?: string;
}> = ({ label, href, value, className }) => (
    <Stack className={`${styles.tagDetailField} ${className ?? ""}`}>
        <Label>{label}</Label>
        <TagLink href={href}>{value}</TagLink>
    </Stack>
);

const TagSection: React.FC<{ title: string; count: number; children: React.ReactNode }> = ({ title, count, children }) => (
    <Stack tokens={{ childrenGap: 10 }} className={styles.detailCard}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Text variant="large">{title}</Text>
            <Text styles={{ root: { color: "#605e5c" } }}>({count})</Text>
        </Stack>
        {count ? children : <Text styles={emptyTextStyles}>No tags selected.</Text>}
    </Stack>
);

const OppNetTagCard: React.FC<{ tag: IOpportunityItem }> = ({ tag }) => (
    <div className={styles.tagDetailCard}>
        <div className={`${styles.tagDetailGrid} ${styles.tagOppNetGrid}`}>
            <LinkFieldDisplay label="Opportunity ID" href={tag.url} value={tag.Id} className={styles.tagDetailIdField} />
            <FieldDisplay label="Title" value={tag.Title} />
            <FieldDisplay label="Customer" value={tag.Customer} />
            <FieldDisplay label="Status" value={tag.Status} />
        </div>
    </div>
);

const PastPerformanceTagCard: React.FC<{ tag: IPastPerformanceItem }> = ({ tag }) => (
    <div className={styles.tagDetailCard}>
        <div className={`${styles.tagDetailGrid} ${styles.tagPastPerformanceGrid}`}>
            <LinkFieldDisplay label="Contract #" href={tag.url} value={tag.Contract_x0023_ || tag.Id} />
            <FieldDisplay label="Customer Agency" value={tag.Customer_x0020_Agency} />
            <FieldDisplay label="Document Type" value={tag.Doc_x0020_Type} />
            <FieldDisplay label="Capability Areas" value={getCapabilityAreas(tag)} />
        </div>
    </div>
);

const ProposalTagCard: React.FC<{ tag: IProposalItem }> = ({ tag }) => (
    <div className={styles.tagDetailCard}>
        <div className={`${styles.tagDetailGrid} ${styles.tagProposalGrid}`}>
            <FieldDisplay label="ID" value={tag.Id} className={styles.tagDetailIdField} />
            <LinkFieldDisplay label="Title" href={tag.url} value={tag.Title || tag.Id} />
            <FieldDisplay label="Opportunity Type" value={tag.TypeOfOpportunity} />
            <FieldDisplay label="Entity" value={tag.Entity} />
            <FieldDisplay label="Stage" value={tag.OpportunityStage} />
        </div>
    </div>
);

export const TaggingInfo: React.FC<ITaggingInfoProps> = ({ capState }) => {
    const oppNetTags = capState.oppNetTags ?? [];
    const pastPerformanceTags = capState.pastPerformanceTags ?? [];
    const proposalTags = capState.proposalTags ?? [];

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <TagSection title="OppNet Opportunities" count={oppNetTags.length}>
                <Stack tokens={{ childrenGap: 10 }}>
                    {oppNetTags.map((tag) => <OppNetTagCard key={`opp-${tag.Id}`} tag={tag} />)}
                </Stack>
            </TagSection>

            <TagSection title="Past Performance" count={pastPerformanceTags.length}>
                <Stack tokens={{ childrenGap: 10 }}>
                    {pastPerformanceTags.map((tag) => <PastPerformanceTagCard key={`past-${tag.Id}`} tag={tag} />)}
                </Stack>
            </TagSection>

            <TagSection title="Proposals" count={proposalTags.length}>
                <Stack tokens={{ childrenGap: 10 }}>
                    {proposalTags.map((tag) => <ProposalTagCard key={`proposal-${tag.Id}`} tag={tag} />)}
                </Stack>
            </TagSection>
        </Stack>
    );
};

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

const FieldDisplay: React.FC<{ label: string; value?: string | number }> = ({ label, value }) => (
    <Stack className={styles.tagDetailField}>
        <Label>{label}</Label>
        <Text>{value || "—"}</Text>
    </Stack>
);

const TagLink: React.FC<{ href?: string; children: React.ReactNode }> = ({ href, children }) =>
    href ? (
        <Link href={href} target="_blank" rel="noopener noreferrer">
            {children}
        </Link>
    ) : (
        <Text>{children}</Text>
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
    <Stack className={styles.tagDetailCard} tokens={{ childrenGap: 8 }}>
        <Stack>
            <Label>Opportunity ID</Label>
            <TagLink href={tag.url}>{tag.Id}</TagLink>
        </Stack>
        <div className={styles.tagDetailGrid}>
            <FieldDisplay label="Title" value={tag.Title} />
            <FieldDisplay label="Customer" value={tag.Customer} />
            <FieldDisplay label="Status" value={tag.Status} />
        </div>
    </Stack>
);

const PastPerformanceTagCard: React.FC<{ tag: IPastPerformanceItem }> = ({ tag }) => (
    <Stack className={styles.tagDetailCard} tokens={{ childrenGap: 8 }}>
        <Stack>
            <Label>Contract #</Label>
            <TagLink href={tag.url}>{tag.Contract_x0023_ || tag.Id}</TagLink>
        </Stack>
        <div className={styles.tagDetailGrid}>
            <FieldDisplay label="Customer Agency" value={tag.Customer_x0020_Agency} />
            <FieldDisplay label="Document Type" value={tag.Doc_x0020_Type} />
            <FieldDisplay label="Capability Areas" value={getCapabilityAreas(tag)} />
        </div>
    </Stack>
);

const ProposalTagCard: React.FC<{ tag: IProposalItem }> = ({ tag }) => (
    <Stack className={styles.tagDetailCard} tokens={{ childrenGap: 8 }}>
        <Stack>
            <Label>Title</Label>
            <TagLink href={tag.url}>{tag.Title || tag.Id}</TagLink>
        </Stack>
        <div className={styles.tagDetailGrid}>
            <FieldDisplay label="ID" value={tag.Id} />
            <FieldDisplay label="Opportunity Type" value={tag.TypeOfOpportunity} />
            <FieldDisplay label="Entity" value={tag.Entity} />
            <FieldDisplay label="Stage" value={tag.OpportunityStage} />
        </div>
    </Stack>
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

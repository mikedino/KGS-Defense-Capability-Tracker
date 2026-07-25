import { IOpportunityItem, IPastPerformanceItem, IProposalItem } from "./props";
import Strings from "./strings";

export const parseJsonTagField = <T>(value?: string): T[] => {
    if (!value?.trim()) return [];

    try {
        const parsed = JSON.parse(value);
        return Array.isArray(parsed) ? parsed as T[] : [];
    } catch {
        return [];
    }
};

export const serializeJsonTagField = <T>(tags?: T[]): string => {
    const cleanTags = tags ?? [];
    return cleanTags.length ? JSON.stringify(cleanTags) : "";
};

const buildListItemUrl = (siteUrl: string, listName: string, itemId: number): string =>
    `${siteUrl}/Lists/${encodeURIComponent(listName)}/DispForm.aspx?ID=${itemId}`;

export const buildOppNetItemUrl = (itemId: number): string =>
    buildListItemUrl(Strings.Sites.oppNet.url, Strings.Sites.oppNet.lists.Opportunities, itemId);

export const buildPastPerformanceItemUrl = (itemId: number): string =>
    buildListItemUrl(Strings.Sites.proposals.url, Strings.Sites.proposals.lists.PastPerformance, itemId);

export const buildProposalItemUrl = (itemId: number): string =>
    buildListItemUrl(Strings.Sites.proposals.url, Strings.Sites.proposals.lists.Proposals, itemId);

export const normalizeOppNetTag = (tag: IOpportunityItem): IOpportunityItem => ({
    Id: tag.Id,
    Title: tag.Title,
    Customer: tag.Customer,
    Status: tag.Status,
    url: tag.url
});

export const normalizePastPerformanceTag = (tag: IPastPerformanceItem): IPastPerformanceItem => ({
    Id: tag.Id,
    Contract_x0023_: tag.Contract_x0023_,
    Customer_x0020_Agency: tag.Customer_x0020_Agency,
    Doc_x0020_Type: tag.Doc_x0020_Type,
    Capability_x0020_Area: tag.Capability_x0020_Area,
    url: tag.url
});

export const normalizeProposalTag = (tag: IProposalItem): IProposalItem => ({
    Id: tag.Id,
    Title: tag.Title,
    OpportunityStage: tag.OpportunityStage,
    TypeOfOpportunity: tag.TypeOfOpportunity,
    Entity: tag.Entity,
    url: tag.url
});

import { IOppNetTagValue } from "./props";

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

export const buildOppNetItemUrl = (oppNetSiteUrl: string, itemId: number): string =>
    `${oppNetSiteUrl}/Lists/Opportunities/DispForm.aspx?ID=${itemId}`;

export const normalizeOppNetTag = (tag: IOppNetTagValue): IOppNetTagValue => ({
    id: tag.id,
    title: tag.title,
    customer: tag.customer,
    status: tag.status,
    url: tag.url
});

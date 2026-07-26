export type CapRouteTab = "overview" | "supporting" | "tagging" | "contract" | "documentation";

export const tabFromSlug = (slug?: string): CapRouteTab => {
    switch ((slug ?? "").toLowerCase()) {
        case "support":
            return "supporting";
        case "tags":
            return "tagging";
        case "contract":
            return "contract";
        case "docs":
            return "documentation";
        case "ov":
        default:
            return "overview";
    }
};

export const tabToSlug = (tab: CapRouteTab): string => {
    switch (tab) {
        case "supporting":
            return "support";
        case "tagging":
            return "tags";
        case "contract":
            return "contract";
        case "documentation":
            return "docs";
        case "overview":
        default:
            return "ov";
    }
};

export const routes = {
    home: "/",
    dashboard: "/dashboard",
    capabilities: "/capabilities",
    contracts: "/contracts",
    admin: "/admin",
    cap: (capId: number, tab: CapRouteTab = "overview"): string => `/caps/${capId}/${tabToSlug(tab)}`,
    contract: (contractId: number): string => `/contracts/${contractId}`
};

export const getPathParts = (pathname: string): string[] =>
    pathname.split("/").map(part => part.trim()).filter(Boolean);

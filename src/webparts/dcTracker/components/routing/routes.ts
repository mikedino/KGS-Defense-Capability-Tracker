export type AppRouteTab = "overview" | "supporting" | "contract" | "documentation";

export const tabFromSlug = (slug?: string): AppRouteTab => {
    switch ((slug ?? "").toLowerCase()) {
        case "support":
            return "supporting";
        case "contract":
            return "contract";
        case "docs":
            return "documentation";
        case "ov":
        default:
            return "overview";
    }
};

export const tabToSlug = (tab: AppRouteTab): string => {
    switch (tab) {
        case "supporting":
            return "support";
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
    app: (appId: number, tab: AppRouteTab = "overview"): string => `/apps/${appId}/${tabToSlug(tab)}`,
    contract: (contractId: number): string => `/contracts/${contractId}`
};

export const getPathParts = (pathname: string): string[] =>
    pathname.split("/").map(part => part.trim()).filter(Boolean);

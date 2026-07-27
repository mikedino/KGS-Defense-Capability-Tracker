export type CapRouteTab = "overview" | "supporting" | "tagging" | "contract" | "documentation";

export const routes = {
    home: "/",
    dashboard: "/dashboard",
    capabilities: "/capabilities",
    contracts: "/contracts",
    admin: "/admin",
    cap: (capId: number): string => `/caps/${capId}/view`,
    contract: (contractId: number): string => `/contracts/${contractId}`
};

export const getPathParts = (pathname: string): string[] =>
    pathname.split("/").map(part => part.trim()).filter(Boolean);

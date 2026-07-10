declare module "react-router-dom" {
    import * as React from "react";

    export interface Location {
        pathname: string;
        search: string;
        hash: string;
        state?: unknown;
    }

    export interface History {
        push(path: string): void;
        replace(path: string): void;
        goBack(): void;
    }

    export interface HashRouterProps {
        basename?: string;
        hashType?: "slash" | "noslash" | "hashbang";
        children?: React.ReactNode;
    }

    export const HashRouter: React.ComponentType<HashRouterProps>;

    export function useHistory(): History;
    export function useLocation(): Location;
}

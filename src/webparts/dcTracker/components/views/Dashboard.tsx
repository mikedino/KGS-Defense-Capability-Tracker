import * as React from "react";
import { Stack, Text, Toggle } from "@fluentui/react";
import {
    DonutChart, HorizontalBarChartWithAxis, IChartProps, IChartDataPoint, IHorizontalBarChartWithAxisDataPoint
} from "@fluentui/react-charting";

import Strings from "../common/strings";
import styles from "../Dct.module.scss";

import type { ICapabilityItem } from "../common/props";

interface IAppDashboardProps {
    capabilities: ICapabilityItem[];
}

type PlatformRank = {
    name: string;
    count: number;
};

const getHorizontalChartMargins = (width: number): { top: number; right: number; bottom: number; left: number } => ({
    top: 20,
    right: width < 480 ? 20 : 40,
    bottom: 40,
    left: width < 480 ? 150 : 220
});
const flexibleChartCellStyle: React.CSSProperties = {
    minWidth: 0,
    maxWidth: "100%"
};
const chartPalette = [
    "#0078d4",
    "#107c10",
    "#8764b8",
    "#ca5010",
    "#038387",
    "#a4262c",
    "#498205",
    "#881798",
    "#004e8c",
    "#8a8886",
    "#c239b3",
    "#ffaa44",
    "#00ad56",
    "#5c2d91"
];

const safeKey = (v: string | undefined): string => (v ?? "").trim() || "Unknown";

const splitAxisLabel = (value: string, maxLineLength = 22, maxLines = 2): string[] => {
    const words = value.split(/\s+/).filter(Boolean);
    if (!words.length) return [value];

    const lines: string[] = [];
    let current = "";

    for (const word of words) {
        const next = current ? `${current} ${word}` : word;

        if (next.length <= maxLineLength || !current) {
            current = next;
        } else {
            lines.push(current);
            current = word;
        }
    }

    if (current) {
        lines.push(current);
    }

    if (lines.length <= maxLines) {
        return lines;
    }

    const visible = lines.slice(0, maxLines);
    visible[maxLines - 1] = `${visible[maxLines - 1].replace(/\.*$/, "")}...`;
    return visible;
};

const getYAxisLabelAvailableWidth = (textNode: SVGTextElement): number => {
    const yAxisGroup = textNode.closest('g[id^="yAxisGElement"]');
    const transform = yAxisGroup?.getAttribute("transform") ?? "";
    const translateMatch = /translate\(([-\d.]+)/.exec(transform);
    const axisX = translateMatch ? Number(translateMatch[1]) : 220;
    const labelX = Number(textNode.getAttribute("x") ?? -10);

    return Math.max(90, axisX + labelX - 8);
};

const wrapSvgYAxisLabels = (container?: HTMLDivElement): void => {
    const textNodes = container?.querySelectorAll<SVGTextElement>('g[id^="yAxisGElement"] .tick text');
    if (!textNodes?.length) return;

    textNodes.forEach((textNode) => {
        const label = textNode.getAttribute("data-full-label") || textNode.textContent || "";
        const x = textNode.getAttribute("x") || "0";
        const y = textNode.getAttribute("y") || "0";
        const baseDy = textNode.getAttribute("dy") || "0.32em";

        textNode.setAttribute("data-full-label", label);
        while (textNode.firstChild) {
            textNode.removeChild(textNode.firstChild);
        }
        textNode.textContent = label;

        const availableWidth = getYAxisLabelAvailableWidth(textNode);
        if (textNode.getComputedTextLength() <= availableWidth) {
            return;
        }

        while (textNode.firstChild) {
            textNode.removeChild(textNode.firstChild);
        }

        const maxLineLength = Math.max(18, Math.floor(availableWidth / 7));
        const lines = splitAxisLabel(label, maxLineLength);
        const initialDy = lines.length > 1 ? "-0.2em" : baseDy;

        lines.forEach((line, index) => {
            const tspan = document.createElementNS("http://www.w3.org/2000/svg", "tspan");
            tspan.setAttribute("x", x);
            tspan.setAttribute("y", y);
            tspan.setAttribute("dy", index === 0 ? initialDy : "1.1em");
            tspan.textContent = line;
            textNode.appendChild(tspan);
        });
    });
};

const useElementWidth = (): [React.RefObject<HTMLDivElement>, number] => {
    const ref = React.useRef<HTMLDivElement>(null);
    const [width, setWidth] = React.useState(0);

    React.useEffect(() => {
        const element = ref.current;
        if (!element || typeof window === "undefined") return;

        const browserWindow = window as Window & typeof globalThis & {
            ResizeObserver?: typeof ResizeObserver;
        };

        const updateWidth = (): void => {
            setWidth(Math.floor(element.getBoundingClientRect().width));
        };

        updateWidth();

        if (browserWindow.ResizeObserver) {
            const observer = new browserWindow.ResizeObserver(updateWidth);
            observer.observe(element);
            return () => observer.disconnect();
        }

        browserWindow.addEventListener("resize", updateWidth);
        return () => browserWindow.removeEventListener("resize", updateWidth);
    }, []);

    return [ref, width];
};

const buildCounts = (capabilities: ICapabilityItem[], selector: (a: ICapabilityItem) => string): Record<string, number> => {
    const counts: Record<string, number> = {};
    for (const a of capabilities) {
        const k = selector(a);
        counts[k] = (counts[k] ?? 0) + 1;
    }
    return counts;
};

const toDonutProps = (
    title: string,
    counts: Record<string, number>,
    showPercentages: boolean,
    colorFor?: (k: string) => string
): IChartProps => {
    const total = Object.values(counts).reduce((s, n) => s + n, 0) || 1;

    const points: IChartDataPoint[] = Object.keys(counts)
        .sort((a, b) => a.localeCompare(b))
        .map((k) => {
            const count = counts[k] ?? 0;
            return {
                legend: k,
                data: showPercentages ? (count / total) * 100 : count,
                color: colorFor?.(k)
            };
        });

    return { chartTitle: title, chartData: points };
};

const toRankedCounts = (counts: Record<string, number>, limit: number): PlatformRank[] =>
    Object.keys(counts)
        .map((name) => ({ name, count: counts[name] ?? 0 }))
        .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
        .slice(0, limit);

const colorForIndex = (index: number): string => chartPalette[index % chartPalette.length];

const capStatusColor = (status: string): string => {
    const normalized = status.trim().toLowerCase();
    const exactColors: Record<string, string> = {
        active: Strings.PillStyles.GreenColor,
        enhancing: Strings.PillStyles.PurpleColor,
        "in development": Strings.PillStyles.SPOBlueColor,
        "reqs gathering": Strings.PillStyles.YellowColor,
        backlog: Strings.PillStyles.GrayColor,
        unknown: Strings.PillStyles.GrayColor
    };

    return exactColors[normalized] ?? colorForIndex(Math.abs(normalized.split("").reduce((sum, char) => sum + char.charCodeAt(0), 0)));
};

const toHorizontalCountBars = (
    rankedCounts: PlatformRank[],
    showPercentages: boolean,
    total: number
): IHorizontalBarChartWithAxisDataPoint[] =>
    rankedCounts.map((item, index) => {
        const value = showPercentages && total > 0 ? (item.count / total) * 100 : item.count;
        return {
            x: value,
            y: item.name,
            legend: item.name,
            color: colorForIndex(index),
            xAxisCalloutData: showPercentages ? `${value.toFixed(1)}%` : item.count.toString(),
            yAxisCalloutData: item.name,
            barLabel: showPercentages ? `${value.toFixed(1)}%` : item.count.toString()
        };
    });

export const AppDashboard: React.FC<IAppDashboardProps> = ({ capabilities }) => {
    const [showPercentages, setShowPercentages] = React.useState(false);
    const [platformChartRef, platformChartWidth] = useElementWidth();
    const [hostingChartRef, hostingChartWidth] = useElementWidth();

    // -------- Stats --------
    const totalCapabilities = capabilities.length;
    // -------- Charts --------
    const capStatusCounts = buildCounts(capabilities, a => safeKey(a.capStatus));
    const platformCounts = buildCounts(capabilities, a => safeKey(a.platform));
    const hostingEnvCounts = buildCounts(capabilities, a => safeKey(a.hostingEnv));

    const capStatusChart = toDonutProps(
        showPercentages ? "Capability Status (%)" : "Capability Status (count)",
        capStatusCounts,
        showPercentages,
        capStatusColor
    );

    const platformTop10 = toRankedCounts(platformCounts, 10);
    const platformBar = toHorizontalCountBars(platformTop10, showPercentages, totalCapabilities);
    const hostingTop10 = toRankedCounts(hostingEnvCounts, 10);
    const hostingBar = toHorizontalCountBars(hostingTop10, showPercentages, totalCapabilities);
    const getChartWidth = (width: number): number => Math.max(280, width || 600);

    React.useEffect(() => {
        const frame = window.requestAnimationFrame(() => {
            window.requestAnimationFrame(() => wrapSvgYAxisLabels(platformChartRef.current ?? undefined));
        });
        return () => window.cancelAnimationFrame(frame);
    }, [platformBar, platformChartRef, platformChartWidth, showPercentages]);

    React.useEffect(() => {
        const frame = window.requestAnimationFrame(() => {
            window.requestAnimationFrame(() => wrapSvgYAxisLabels(hostingChartRef.current ?? undefined));
        });
        return () => window.cancelAnimationFrame(frame);
    }, [hostingBar, hostingChartRef, hostingChartWidth, showPercentages]);

    return (
        <Stack
            tokens={{ childrenGap: 24 }}
            className={styles.dashboardContent}
            styles={{ root: { width: "100%", maxWidth: "100%", alignSelf: "stretch" } }}
        >
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack>
                    <Text variant="large" style={{ fontWeight: 500 }}>
                        Capability inventory, usage, and status summary | Total Capabilities: {totalCapabilities.toLocaleString()}
                    </Text>
                </Stack>

                <Toggle
                    label="Show percentages"
                    checked={showPercentages}
                    onChange={(_, checked) => setShowPercentages(!!checked)}
                    inlineLabel
                />
            </Stack>

            {/* Main charts row */}
            <div className={styles.dashboardTwoColumnGrid}>
                <Stack tokens={{ childrenGap: 6 }} horizontalAlign="center" className={styles.dashboardCell}>
                    <Text variant="mediumPlus" style={{ fontWeight: 600, textAlign: "center" }}>Capability Status</Text>
                    <DonutChart
                        data={capStatusChart}
                        innerRadius={45}
                        legendProps={{ canSelectMultipleLegends: false, allowFocusOnLegends: true }}
                        hideLabels={false}
                        showLabelsInPercent={showPercentages}
                        height={250}
                        width={350}
                        roundCorners
                    />
                </Stack>

                {/*
                <Stack tokens={{ childrenGap: 6 }} horizontalAlign="center" className={styles.dashboardCell}>
                    <Text variant="mediumPlus" style={{ fontWeight: 600, textAlign: "center" }}>Platform</Text>
                    <DonutChart
                        data={platformChart}
                        innerRadius={45}
                        legendProps={{ canSelectMultipleLegends: false, allowFocusOnLegends: true }}
                        hideLabels={false}
                        showLabelsInPercent={showPercentages}
                        height={250}
                        width={350}
                        roundCorners
                    />
                </Stack>
                */}

            </div>

            {/* Top bar charts */}
            <div className={styles.dashboardBarGrid}>
                <Stack tokens={{ childrenGap: 8 }} className={styles.dashboardCell} style={flexibleChartCellStyle}>
                    <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
                        Top 10 Platforms {showPercentages ? "(%)" : "(count)"}
                    </Text>
                    <div ref={platformChartRef} className={styles.dashboardChartFrame}>
                        <HorizontalBarChartWithAxis
                            data={platformBar}
                            height={360}
                            width={getChartWidth(platformChartWidth)}
                            barHeight={20}
                            hideLegend={false}
                            margins={getHorizontalChartMargins(platformChartWidth)}
                            roundCorners
                        />
                    </div>
                </Stack>

                <Stack tokens={{ childrenGap: 8 }} className={styles.dashboardCell} style={flexibleChartCellStyle}>
                    <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
                        Top 10 Hosting Environments {showPercentages ? "(%)" : "(count)"}
                    </Text>
                    <div ref={hostingChartRef} className={styles.dashboardChartFrame}>
                        <HorizontalBarChartWithAxis
                            data={hostingBar}
                            height={360}
                            width={getChartWidth(hostingChartWidth)}
                            barHeight={20}
                            hideLegend={false}
                            margins={getHorizontalChartMargins(hostingChartWidth)}
                            roundCorners
                        />
                    </div>
                </Stack>
            </div>

            {/* Optional second donut row ideas (enable later):
          - Environment donut (envCounts)
          - Connectivity donut (connectivityCounts)
      */}
        </Stack>
    );
};

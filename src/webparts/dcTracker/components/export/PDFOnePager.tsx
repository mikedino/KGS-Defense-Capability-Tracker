import * as React from "react";
import { Font, Page, Text, View, Image, Link, StyleSheet } from "@react-pdf/renderer";
import { ICapabilityItem, IContractItem } from "../common/props";
import Strings from "../common/strings";
//import { renderRichText } from "./HTMLParser";
import { formatDate } from "../common/utils";

export interface ICapabilityOnePagerProps {
    capability: ICapabilityItem;
    contract?: IContractItem;
    kgsLogoDataUrl?: string;
    screenshotBinary?: string;
}

// Disable automatic hyphenation globally for this PDF renderer
Font.registerHyphenationCallback((word) => [word]);

const styles = StyleSheet.create({
    page: {
        paddingTop: 34,
        paddingHorizontal: 28,
        paddingBottom: 30,
        fontSize: 10,
        fontFamily: "Helvetica",
        lineHeight: 1.25
    },

    header: {
        flexDirection: "row",
        alignItems: "center"
    },
    logo: { width: 150, height: 50, objectFit: "contain" },
    headerTitleWrap: { flex: 1, paddingRight: 12 },
    title: { fontSize: 18, fontWeight: 700, lineHeight: 1.05 },
    subtitle: { marginTop: 6, fontSize: 10 },

    // header: {
    //     flexDirection: "row",
    //     justifyContent: "space-between",
    //     alignItems: "flex-start",
    //     width: "100%",
    //     marginBottom: 4,
    //     minHeight: 0
    // },

    // headerTitleWrap: { width: "72%", paddingRight: 12, minHeight: 0, flexShrink: 1 },
    // logoWrap: { width: "28%", alignItems: "flex-end", alignContent: "flex-start" },
    // logo: { width: 150, height: 50, objectFit: "contain" },
    // title: { fontSize: 16, fontWeight: 700 },
    // subtitle: { marginTop: 4, fontSize: 10 },

    rule: { marginTop: 0, height: 3, backgroundColor: Strings.PillStyles.Merlot },

    body: { marginTop: 12, flexDirection: "row", gap: 12 },
    leftCol: { flex: 1.25 },
    rightCol: { flex: 0.85 },

    sectionTitle: { fontSize: 11, fontWeight: 700, marginBottom: 4 },
    para: { marginBottom: 12 },

    labelRow: { flexDirection: "row", marginBottom: 3 },
    label: { width: 110, fontWeight: 700 },
    value: { flex: 1, paddingRight: 6 },

    cardTitle: {
        fontSize: 11,
        fontWeight: 700,
        marginBottom: 6,
        color: "#FFFFFF"
    },

    kpiCard: {
        backgroundColor: Strings.PillStyles.Matisse,
        color: "#FFFFFF",
        borderRadius: 12,
        padding: 10,
        marginBottom: 10
    },

    technicalInfoCard: {
        backgroundColor: Strings.PillStyles.Tarawera,
        color: "#FFFFFF",
        borderRadius: 12,
        padding: 10,
        marginBottom: 10
    },

    listItem: {
        flexDirection: "row",
        marginBottom: 4
    },
    bullet: {
        width: 10,
        fontWeight: 700
    },

    screenshotWrap: {
        flexDirection: "row",
        marginTop: 8,
        borderWidth: 1,
        borderColor: "#bbb",
        borderRadius: 6,
        padding: 6
    },
    screenshot: {
        width: "100%",
        height: 210,
        objectFit: "contain"
    },

    classificationHeader: {
        position: "absolute",
        top: 6,
        left: 0,
        right: 0,
        textAlign: "center",
        fontSize: 10,
        fontWeight: 700,
        letterSpacing: 1
    },

    footer: {
        position: "absolute",
        left: 28,
        right: 28,
        bottom: 12,
        fontSize: 8,

        color: "#666",
        flexDirection: "row",
        alignItems: "center"
    },

    footerLeft: {
        flex: 1,
        textAlign: "left",
        overflow: "hidden"
    },

    footerCenter: {
        flex: 1,
        textAlign: "center",
        fontWeight: 700,
        letterSpacing: 1
    },

    footerRight: {
        flex: 1,
        textAlign: "right"
    },

    link: {
        color: "#0563C1",
        textDecoration: "underline",
        wordBreak: "break-all"
    }
});

const shortTitleFooter = (title: string): string => {
    return title.length > 50 ? title.substring(0, 47) + "..." : title;
}

const renderBulletList = (items: string[], textColor?: string): JSX.Element => {
    const colorStyle = textColor ? { color: textColor } : {};

    return (
        <View>
            {items.length > 0 ? (
                items.map((item, index) => (
                    <View key={`${item}-${index}`} style={styles.listItem}>
                        <Text style={[styles.bullet, colorStyle]}>•</Text>
                        <Text style={[{ flex: 1 }, colorStyle]}>{item}</Text>
                    </View>
                ))
            ) : (
                <Text style={colorStyle}>—</Text>
            )}
        </View>
    );
};

const CapabilityOnePager: React.FC<ICapabilityOnePagerProps> = (props) => {
    const { capability, contract, kgsLogoDataUrl, screenshotBinary } = props;

    const description = (capability.description ?? "").trim();

    const technicalInfoList: string[] = [
        capability.capStatus ? `Capability Status: ${capability.capStatus}` : "",
        capability.platform ? `Platform: ${capability.platform}` : "",
        capability.hostingEnv ? `Hosting Environment: ${capability.hostingEnv}` : "",
        capability.connectivity ? `Connectivity: ${capability.connectivity}` : "",
        capability.compliance ? `Compliance: ${capability.compliance}` : "",
        capability.licenseReqd ? `License Required? ${capability.licenseReqd}` : "",
        capability.extensibility ? `APIs/Extensibility: ${capability.extensibility}` : "",
        capability.codeLanguage ? `Coding Language: ${capability.codeLanguage}` : ""
    ].filter(Boolean);

    return (
        <Page size="LETTER" orientation="portrait" style={styles.page} wrap={false}>
            {/* <Text style={styles.classificationHeader} fixed>
                UNCLASSIFIED
            </Text> */}

            <View style={styles.header}>
                <View style={styles.headerTitleWrap}>
                    <Text style={styles.title}>{capability.Title}</Text>
                    {/* <Text style={styles.subtitle}>Capability One-Page Summary</Text> */}
                </View>

                {kgsLogoDataUrl ? (
                    <Image style={styles.logo} src={kgsLogoDataUrl} />
                ) : (
                    <View style={styles.logo} />
                )}
            </View>

            <View style={styles.rule} />

            <View style={styles.body}>
                <View style={styles.leftCol}>
                    <Text style={styles.sectionTitle}>Description</Text>
                    <Text style={styles.para}>{description || "—"}</Text>

                    <Text style={styles.sectionTitle}>General Info</Text>

                    <View style={[styles.para, { marginTop: 4 }]}>
                        <View style={styles.labelRow}>
                            <Text style={styles.label}>Status</Text>
                            <Text style={styles.value}>{capability.capStatus || "—"}</Text>
                        </View>

                        <View style={styles.labelRow}>
                            <Text style={styles.label}>URL</Text>
                            <View style={styles.value}>
                                {capability.link ? (
                                    <Link src={capability.link} style={styles.link}>
                                        Capability URL
                                    </Link>
                                ) : (
                                    <Text>—</Text>
                                )}
                            </View>
                        </View>

                        <View style={styles.labelRow}>
                            <Text style={styles.label}>Contract</Text>
                            <Text style={styles.value}>{contract?.Title ?? "—"}</Text>
                        </View>
                    </View>

                    <Text style={styles.sectionTitle}>Technical Capabilities</Text>
                    <Text style={styles.para}>{capability.capabilities || "—"}</Text>
                </View>

                <View style={styles.rightCol}>
                    {/* <View style={styles.kpiCard}>
                        <Text style={styles.cardTitle}>Highlights</Text>
                        {renderRichText(capability.highlights, { color: "#FFFFFF" })}
                    </View> */}

                    <View style={styles.technicalInfoCard}>
                        <Text style={styles.cardTitle}>Technical Info</Text>
                        {renderBulletList(technicalInfoList, "#FFFFFF")}
                    </View>
                </View>
            </View>

            <View style={styles.screenshotWrap}>
                {screenshotBinary ? (
                    <Image style={styles.screenshot} src={screenshotBinary} />
                ) : (
                    <Text>No screenshot available.</Text>
                )}
            </View>

            <View style={styles.footer}>
                <Text style={styles.footerLeft}>{shortTitleFooter(capability.Title)}</Text>
                <Text style={styles.footerCenter}>UNCLASSIFIED</Text>
                <Text style={styles.footerRight}>{formatDate(capability.Modified ?? "")}</Text>
            </View>
        </Page>
    );
};

export default CapabilityOnePager;

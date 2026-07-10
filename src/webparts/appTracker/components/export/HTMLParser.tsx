/**
 * 
The HTML formatting is not automatically rendered by React-PDF.
So we need to explicitly translate tags like <strong> and <em> into nested <Text> styles.

This parser: 

keeps <strong> as bold
keeps <em> / <i> as italic
keeps <u> as underline
preserves <p> and <div> as paragraph blocks
renders <ul> / <ol> lists properly
handles <br>
can also render links if SharePoint outputs them
 */

import * as React from "react";
import { Text, View, Link } from "@react-pdf/renderer";

type RichTextStyle = {
    color?: string;
};

type InlineStyle = {
    color?: string;
    fontWeight?: "bold" | "normal" | number;
    fontStyle?: "italic" | "normal";
    textDecoration?: "underline" | "none";
};

const decodeHtml = (value: string): string =>
    value
        .replace(/&nbsp;/g, " ")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, "\"")
        .replace(/&#39;/g, "'");

const normalizeHtml = (html?: string): string =>
    (html ?? "").trim();

const mergeInlineStyles = (
    base?: InlineStyle,
    extra?: InlineStyle
): InlineStyle => ({
    ...(base ?? {}),
    ...(extra ?? {})
});

const getElementChildren = (node: Node): Node[] =>
    Array.from(node.childNodes ?? []);

const renderInlineNodes = (
    nodes: Node[],
    inheritedStyle?: InlineStyle,
    keyPrefix: string = "inline"
): React.ReactNode[] => {
    return nodes.map((node, index) => {
        const key = `${keyPrefix}-${index}`;

        if (node.nodeType === Node.TEXT_NODE) {
            const textValue = decodeHtml(node.textContent ?? "");
            return textValue ? (
                <Text key={key} style={inheritedStyle ?? {}}>
                    {textValue}
                </Text>
            ) : null;
        }

        if (node.nodeType !== Node.ELEMENT_NODE) {
            return null;
        }

        const element = node as HTMLElement;
        const tag = element.tagName.toLowerCase();

        if (tag === "br") {
            return (
                <Text key={key} style={inheritedStyle ?? {}}>
                    {"\n"}
                </Text>
            );
        }

        if (tag === "strong" || tag === "b") {
            const nextStyle = mergeInlineStyles(inheritedStyle, { fontWeight: "bold" });
            return (
                <Text key={key} style={nextStyle}>
                    {renderInlineNodes(getElementChildren(element), nextStyle, key)}
                </Text>
            );
        }

        if (tag === "em" || tag === "i") {
            const nextStyle = mergeInlineStyles(inheritedStyle, { fontStyle: "italic" });
            return (
                <Text key={key} style={nextStyle}>
                    {renderInlineNodes(getElementChildren(element), nextStyle, key)}
                </Text>
            );
        }

        if (tag === "u") {
            const nextStyle = mergeInlineStyles(inheritedStyle, { textDecoration: "underline" });
            return (
                <Text key={key} style={nextStyle}>
                    {renderInlineNodes(getElementChildren(element), nextStyle, key)}
                </Text>
            );
        }

        if (tag === "a") {
            const href = element.getAttribute("href") ?? "";
            const linkStyle = mergeInlineStyles(inheritedStyle, { textDecoration: "underline" });

            return href ? (
                <Link key={key} src={href} style={linkStyle}>
                    {renderInlineNodes(getElementChildren(element), linkStyle, key)}
                </Link>
            ) : (
                <Text key={key} style={inheritedStyle ?? {}}>
                    {renderInlineNodes(getElementChildren(element), inheritedStyle, key)}
                </Text>
            );
        }

        return (
            <Text key={key} style={inheritedStyle ?? {}}>
                {renderInlineNodes(getElementChildren(element), inheritedStyle, key)}
            </Text>
        );
    });
};

const renderBlockNode = (
    node: Node,
    textStyle?: RichTextStyle,
    key: string = "block"
): React.ReactNode => {
    if (node.nodeType === Node.TEXT_NODE) {
        const textValue = decodeHtml(node.textContent ?? "").trim();
        return textValue ? (
            <Text key={key} style={[{ marginBottom: 6 }, textStyle ?? {}]}>
                {textValue}
            </Text>
        ) : null;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
        return null;
    }

    const element = node as HTMLElement;
    const tag = element.tagName.toLowerCase();
    const inlineBaseStyle: InlineStyle = textStyle ?? {};

    if (tag === "p" || tag === "div") {
        return (
            <Text key={key} style={[{ marginBottom: 6 }, textStyle ?? {}]}>
                {renderInlineNodes(getElementChildren(element), inlineBaseStyle, key)}
            </Text>
        );
    }

    if (tag === "ul") {
        const items = Array.from(element.children).filter(
            (child) => child.tagName.toLowerCase() === "li"
        );

        return (
            <View key={key} style={{ marginBottom: 6 }}>
                {items.map((item, index) => (
                    <View
                        key={`${key}-li-${index}`}
                        style={{ flexDirection: "row", marginBottom: 3 }}
                    >
                        <Text style={[{ width: 10 }, textStyle ?? {}]}>•</Text>
                        <Text style={[{ flex: 1 }, textStyle ?? {}]}>
                            {renderInlineNodes(getElementChildren(item), inlineBaseStyle, `${key}-li-${index}`)}
                        </Text>
                    </View>
                ))}
            </View>
        );
    }

    if (tag === "ol") {
        const items = Array.from(element.children).filter(
            (child) => child.tagName.toLowerCase() === "li"
        );

        return (
            <View key={key} style={{ marginBottom: 6 }}>
                {items.map((item, index) => (
                    <View
                        key={`${key}-li-${index}`}
                        style={{ flexDirection: "row", marginBottom: 3 }}
                    >
                        <Text style={[{ width: 18 }, textStyle ?? {}]}>{`${index + 1}.`}</Text>
                        <Text style={[{ flex: 1 }, textStyle ?? {}]}>
                            {renderInlineNodes(getElementChildren(item), inlineBaseStyle, `${key}-li-${index}`)}
                        </Text>
                    </View>
                ))}
            </View>
        );
    }

    if (tag === "br") {
        return (
            <Text key={key} style={textStyle ?? {}}>
                {"\n"}
            </Text>
        );
    }

    return (
        <Text key={key} style={[{ marginBottom: 6 }, textStyle ?? {}]}>
            {renderInlineNodes(getElementChildren(element), inlineBaseStyle, key)}
        </Text>
    );
};

export const renderRichText = (html?: string, textStyle?: RichTextStyle): JSX.Element => {
    const normalized = normalizeHtml(html);

    if (!normalized) {
        return <Text style={textStyle ?? {}}>—</Text>;
    }

    if (typeof DOMParser === "undefined") {
        const fallback = decodeHtml(
            normalized
                .replace(/<br\s*\/?>/gi, "\n")
                .replace(/<\/p>/gi, "\n\n")
                .replace(/<[^>]+>/g, "")
        ).trim();

        return <Text style={textStyle ?? {}}>{fallback || "—"}</Text>;
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(normalized, "text/html");
    const bodyNodes = Array.from(doc.body.childNodes);

    const content = bodyNodes
        .map((node, index) => renderBlockNode(node, textStyle, `block-${index}`))
        .filter(Boolean);

    return <View>{content.length > 0 ? content : <Text style={textStyle ?? {}}>—</Text>}</View>;
};
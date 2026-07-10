import * as React from "react";
import { RichText, StyleOptions } from "@pnp/spfx-controls-react/lib/RichText";
import Strings from "../../strings";
//import { useTheme } from "@fluentui/react";
import { Label } from "@fluentui/react";

export interface IRichTextFieldProps {
    label: string;
    value: string;
    maxLength: number;
    description?: string;
    errorMessage?: string;
    className?: string;
    required?: boolean;
    placeholder?: string;
    onChange: (value: string) => string;
}

const basicStyleOptions: StyleOptions = {
    showAlign: false,
    showBold: true,
    showItalic: true,
    showLink: false,
    showList: true,
    showMore: false,
    showImage: false,
    showStyles: false,
    showUnderline: false
};

export const stripHtml = (html: string): string => {
    const temp: HTMLDivElement = document.createElement("div");
    temp.innerHTML = html;

    return (temp.textContent || temp.innerText || "")
        .replace(/\u00A0/g, " ")
        .replace(/\s+/g, " ")
        .trim();
};

export const getPlainTextLength = (html: string): number => {
    return stripHtml(html).length;
};

export const RichTextField = ({
    label,
    value,
    maxLength,
    description,
    errorMessage,
    className,
    required = false,
    placeholder,
    onChange
}: IRichTextFieldProps): JSX.Element => {
    //const theme = useTheme();

    const plainTextLength: number = getPlainTextLength(value);
    const isOverLimit: boolean = plainTextLength > maxLength;

    const computedErrorMessage: string | undefined = errorMessage
        ?? (isOverLimit ? `Please keep to ${maxLength} characters or fewer.` : undefined);

    const handleEditorChange = (newValue: string): string => {
        return onChange(newValue);
    };

    const handleBeforePaste = (event: React.ClipboardEvent<HTMLDivElement>): boolean => {
        const clipboardText: string = event.clipboardData.getData("text/plain") ?? "";
        const currentText: string = stripHtml(value);
        const selectionText: string = window.getSelection()?.toString() ?? "";

        const nextLength: number = currentText.length - selectionText.length + clipboardText.length;

        if (nextLength <= maxLength) {
            return true;
        }

        event.preventDefault();

        const remaining: number = Math.max(0, maxLength - (currentText.length - selectionText.length));
        const truncatedText: string = clipboardText.slice(0, remaining);

        if (truncatedText) {
            const escapedText: string = truncatedText
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#39;");

            document.execCommand("insertHTML", false, escapedText);
        }

        return false;
    };

    // const fontStyle: React.CSSProperties = {
    //     ["--rtf-font-size" as keyof React.CSSProperties]: theme.fonts.medium.fontSize,
    //     ["--rtf-line-height" as keyof React.CSSProperties]: theme.fonts.medium.lineHeight,
    //     ["--rtf-font-family" as keyof React.CSSProperties]: theme.fonts.medium.fontFamily,
    //     ["--rtf-font-weight" as keyof React.CSSProperties]: String(theme.fonts.medium.fontWeight)
    // };

    return (
        <div className={className} style={{ margin: 6 }}>
            <Label required={required}>{label}</Label>

            {description && (
                <div className="ms-TextField-description">
                    {description}
                </div>
            )}

            <div onPaste={handleBeforePaste}>
                <RichText
                    value={value}
                    isEditMode={true}
                    onChange={handleEditorChange}
                    styleOptions={basicStyleOptions}
                    placeholder={placeholder}                    
                />
            </div>

            <div
                className="ms-TextField-description"
                style={isOverLimit ? { color: Strings.PillStyles.RedColor } : undefined}
            >
                {plainTextLength}/{maxLength}
            </div>

            {computedErrorMessage && (
                <div
                    className="ms-TextField-errorMessage"
                    style={{ display: "block", color: Strings.PillStyles.RedColor }}
                >
                    {computedErrorMessage}
                </div>
            )}
        </div>
    );
};
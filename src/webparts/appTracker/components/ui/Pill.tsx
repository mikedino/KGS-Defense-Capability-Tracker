import * as React from "react";

export interface PillProps {
  text: string;
  backgroundColor?: string;
  textColor?: string;
  style?: React.CSSProperties;
}

export const Pill: React.FC<PillProps> = ({
  text,
  backgroundColor = "#e0e0e0",
  textColor = "#000",
  style,
}) => {
  return (
    <span
      style={{
        backgroundColor,
        color: textColor,
        padding: "4px 10px",
        borderRadius: "20px",
        display: "inline-block",
        //fontWeight: 600,
        fontSize: "0.75rem",
        textTransform: "capitalize",
        ...style,
      }}
    >
      {text}
    </span>
  );
};

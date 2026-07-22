import * as React from "react";
import { Stack, Text, Icon } from "@fluentui/react";

export interface StatCardProps {
  title: string;
  value: string;
  iconName: string;
  color: string;
}

export const StatCard: React.FC<StatCardProps> = ({ title, value, iconName, color }) => {

    return (
    <div
      style={{
        backgroundColor: "#f3f2f1",
        padding: "16px 20px",
        borderRadius: "4px",
        borderLeft: `4px solid ${color}`,
        minWidth: "140px",
        textAlign: "center",
      }}
    >
      <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }}>
        <Icon iconName={iconName} style={{ fontSize: "20px", color }} />
        <Text variant="xLarge" style={{ fontWeight: 600, color }}>
          {value}
        </Text>
        <Text variant="medium">{title}</Text>
      </Stack>
    </div>
    )
    
}
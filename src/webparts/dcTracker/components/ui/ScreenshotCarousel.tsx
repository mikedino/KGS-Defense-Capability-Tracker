import React, { useState } from "react";
import { Stack, Icon, Text } from "@fluentui/react";
import styles from "../Dct.module.scss";

interface Screenshot {
  UniqueId: string;
}

interface Props {
  screenshots: Screenshot[];
  solutionTitle: string;
  webUrl: string;
}

export const ScreenshotCarousel: React.FC<Props> = ({ screenshots, solutionTitle, webUrl }) => {
  const [currentIndex, setCurrentIndex] = useState(0);

  if (!screenshots || screenshots.length === 0) return null;

  const currentScreenshot = screenshots[currentIndex];

  const goPrev = () : void => {
    setCurrentIndex((prev) => (prev === 0 ? screenshots.length - 1 : prev - 1));
  };

  const goNext = () : void => {
    setCurrentIndex((prev) => (prev === screenshots.length - 1 ? 0 : prev + 1));
  };

  return (
    <Stack>
      <Stack className={styles.screenshot} horizontal horizontalAlign="center" verticalAlign="center">
        {screenshots.length > 1 && (
          <Icon
            iconName="ChevronLeft"
            onClick={goPrev}
            style={{ fontSize: 24, cursor: "pointer", marginRight: 8 }}
          />
        )}

        <iframe
          src={`${webUrl}/_layouts/15/embed.aspx?UniqueId=${currentScreenshot.UniqueId}`}
          title={`${solutionTitle} Screenshot`}
          className="img-fluid img-thumbnail"
        />

        {screenshots.length > 1 && (
          <Icon
            iconName="ChevronRight"
            onClick={goNext}
            style={{ fontSize: 24, cursor: "pointer", marginLeft: 8 }}
          />
        )}
      </Stack>

      <Stack horizontalAlign="center" horizontal tokens={{ childrenGap: 6 }}>
        <Text variant="small">
          Screenshot {currentIndex + 1} of {screenshots.length}
        </Text>
        <Icon
          iconName="ZoomIn"
          title="Open the full image in a new window"
          className="printHide"
          onClick={() => {
            window.open(`${webUrl}/_layouts/15/embed.aspx?UniqueId=${currentScreenshot.UniqueId}`, "_blank");
          }}
          style={{ fontSize: 14, cursor: "pointer" }}
        />
      </Stack>
    </Stack>
  );
};

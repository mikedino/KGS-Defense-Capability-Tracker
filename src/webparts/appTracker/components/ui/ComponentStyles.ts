import { IPivotStyles, IButtonStyles, IStackStyles } from "@fluentui/react";
import Strings from "../../strings";

/***** Custom tabs (PIVOTS) *****/
export const customPivotStyles: Partial<IPivotStyles> = {
  link: {
    backgroundColor: Strings.PillStyles.GrayFill,
    color: "#333",
    selectors: {
      "&.is-selected": {
        backgroundColor: Strings.PillStyles.Matisse,
        borderBottomColor: Strings.PillStyles.Matisse,
        color: "white",
      },
      "&.is-selected:hover, &:hover.is-selected": { // selected + hover state
        backgroundColor: Strings.PillStyles.Matisse,
        borderBottomColor: Strings.PillStyles.Matisse,
        color: "white",
      },
      "&:hover": {
        backgroundColor: Strings.PillStyles.NavHoverFill,
        color: "#323130",
      }
    }
  }
};

/***** App, Flow, Iteration card galleries *****/
export const compactButtonStyles: IButtonStyles = {
  root: {
    height: 24,
    minHeight: 24,
    padding: "0 6px",
    fontSize: 12,
    fontWeight: 400,
    lineHeight: "20px",
    borderRadius: 4,
    border: "1px solid #8f8c8a"
  },
  rootHovered: {
    borderColor: "#0078d4", // border on hover
  },
  label: {
    fontSize: 12,
    fontWeight: 400,
  },
  icon: {
    fontSize: 12,
    color: "#0078d4"
  },
};

/***** Forms sections *****/
export const cardStackStyles: IStackStyles = {
  root: {
    padding: 16,
    borderRadius: 8,
    boxShadow: "inset 0px 0px 7px 5px rgba(0, 0, 0, .1)",
    background: "#fdfdfd",
    border: `1px solid #aeaeae` // was #edebe9
  },
};

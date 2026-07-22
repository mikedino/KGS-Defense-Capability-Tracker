import Strings from "../common/strings";

interface ICapStatusColors {
  backgroundColor: string;
  textColor: string;
}

const DefaultStatusColors: ICapStatusColors = {
  backgroundColor: Strings.PillStyles.GrayFill,
  textColor: Strings.PillStyles.GrayColor
};

// ----------------------
// Capability Status - KNOWN statuses
// ----------------------
const CapStatusColorMap: Partial<Record<string, ICapStatusColors>> = {
  Active: {
    backgroundColor: Strings.PillStyles.GreenFill,
    textColor: Strings.PillStyles.GreenColor
  },
  "In Development": {
    backgroundColor: Strings.PillStyles.BlueFill,
    textColor: Strings.PillStyles.BlueColor
  },
  Enhancing: {
    backgroundColor: Strings.PillStyles.PurpleFill,
    textColor: Strings.PillStyles.PurpleColor
  },
  "Reqs Gathering": {
    backgroundColor: Strings.PillStyles.YellowFill,
    textColor: Strings.PillStyles.YellowColor
  },
  Backlog: {
    backgroundColor: Strings.PillStyles.GrayFill,
    textColor: Strings.PillStyles.GrayColor
  }
};

export const getAtoStatusFill = (status?: string): ICapStatusColors => {
  if (!status) return DefaultStatusColors;

  const key = status.trim();
  return CapStatusColorMap[key] ?? DefaultStatusColors;
};

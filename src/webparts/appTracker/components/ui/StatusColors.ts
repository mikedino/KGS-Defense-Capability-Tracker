import Strings from "../../strings";

interface IAppStatusColors {
  backgroundColor: string;
  textColor: string;
}

const DefaultStatusColors: IAppStatusColors = {
  backgroundColor: Strings.PillStyles.GrayFill,
  textColor: Strings.PillStyles.GrayColor
};

// ----------------------
// App Status - KNOWN statuses
// ----------------------
const AppStatusColorMap: Partial<Record<string, IAppStatusColors>> = {
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

export const getAtoStatusFill = (status?: string): IAppStatusColors => {
  if (!status) return DefaultStatusColors;

  const key = status.trim();
  return AppStatusColorMap[key] ?? DefaultStatusColors;
};

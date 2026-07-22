import { createTheme } from "@fluentui/react";

export const appTheme = createTheme({
  components: {
    Label: {
      styles: {
        root: {
          color: "#205493"
        }
      }
    },
    DatePicker: {
      styles: {
        root: {
          margin: "6px"
        }
      }
    },
    TextField: {
      styles: {
        root: {
          margin: "6px"
        }
      }
    }
  }
});
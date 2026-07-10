import * as React from "react";
import { Document } from "@react-pdf/renderer";
import { IApplicationItem, IContractItem } from "../data/props";
import ApplicationOnePager from "./PDFOnePager";

export interface IApplicationBookItem {
  application: IApplicationItem;
  contract?: IContractItem;
  screenshotBinary?: string;
}

export interface IApplicationOnePagerBookProps {
  items: IApplicationBookItem[];
  kgsLogoDataUrl?: string;
}

const ApplicationBook: React.FC<IApplicationOnePagerBookProps> = (props) => {
  const { items, kgsLogoDataUrl } = props;

  return (
    <Document>
      {items.map((i) => (
        <ApplicationOnePager
          key={i.application.Id}
          application={i.application}
          contract={i.contract}
          kgsLogoDataUrl={kgsLogoDataUrl}
          screenshotBinary={i.screenshotBinary}
        />
      ))}
    </Document>
  );
};

export default ApplicationBook;
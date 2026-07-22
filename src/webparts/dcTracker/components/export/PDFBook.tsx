import * as React from "react";
import { Document } from "@react-pdf/renderer";
import { ICapabilityItem, IContractItem } from "../common/props";
import CapabilityOnePager from "./PDFOnePager";

export interface ICapabilityBookItem {
  capability: ICapabilityItem;
  contract?: IContractItem;
  screenshotBinary?: string;
}

export interface ICapabilityOnePagerBookProps {
  items: ICapabilityBookItem[];
  kgsLogoDataUrl?: string;
}

const CapabilityBook: React.FC<ICapabilityOnePagerBookProps> = (props) => {
  const { items, kgsLogoDataUrl } = props;

  return (
    <Document>
      {items.map((i) => (
        <CapabilityOnePager
          key={i.capability.Id}
          capability={i.capability}
          contract={i.contract}
          kgsLogoDataUrl={kgsLogoDataUrl}
          screenshotBinary={i.screenshotBinary}
        />
      ))}
    </Document>
  );
};

export default CapabilityBook;
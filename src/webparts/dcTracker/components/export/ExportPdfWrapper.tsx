import * as React from "react";
import { Document, pdf } from "@react-pdf/renderer";
import { saveAs } from "file-saver";
import { ICapabilityItem, IContractItem } from "../common/props";
import CapabilityOnePager from "./PDFOnePager";
import CapabilityBook, { ICapabilityBookItem } from "./PDFBook";

export interface IExportCapabilityPdfArgs {
  capability: ICapabilityItem;
  contract?: IContractItem;
  kgsLogoDataUrl?: string;
  screenshotBinary?: string;
  fileName?: string;
}

export interface IExportCapabilitiesBookArgs {
  items: ICapabilityBookItem[];      // filtered capabilities mapped to export payload
  kgsLogoDataUrl?: string;
  fileName?: string;
}

export const exportCapabilityPdf = async (args: IExportCapabilityPdfArgs): Promise<void> => {

  const baseName: string =
    (args.fileName ?? args.capability.Title)
      .replace(/\s+/g, "")   // remove ALL spaces
      .replace(/[\\/:*?"<>|]+/g, "-");  // sanitize illegal characters

  const safeFileName: string = `${baseName}_FactSheet.pdf`;

  const doc: React.ReactElement = (
    <Document>
      <CapabilityOnePager
        capability={args.capability}
        contract={args.contract}
        kgsLogoDataUrl={args.kgsLogoDataUrl}
        screenshotBinary={args.screenshotBinary}
      />
    </Document>
  );

  const blob: Blob = await pdf(doc).toBlob();
  saveAs(blob, safeFileName);

};


export const exportCapabilitiesBookPdf = async (
  args: IExportCapabilitiesBookArgs
): Promise<void> => {

  const baseName: string = (args.fileName ?? "Capabilities")
    .replace(/\s+/g, "")
    .replace(/[\\/:*?"<>|]+/g, "-");

  const safeFileName: string = `${baseName}_SummaryBook.pdf`;

  const doc: React.ReactElement = (
    <CapabilityBook
      items={args.items}
      kgsLogoDataUrl={args.kgsLogoDataUrl}
    />
  );

  const blob: Blob = await pdf(doc).toBlob();
  saveAs(blob, safeFileName);
};
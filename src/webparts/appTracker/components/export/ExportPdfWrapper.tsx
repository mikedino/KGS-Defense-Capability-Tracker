import * as React from "react";
import { Document, pdf } from "@react-pdf/renderer";
import { saveAs } from "file-saver";
import { IApplicationItem, IContractItem } from "../data/props";
import ApplicationOnePager from "./PDFOnePager";
import ApplicationBook, { IApplicationBookItem } from "./PDFBook";

export interface IExportApplicationPdfArgs {
  application: IApplicationItem;
  contract?: IContractItem;
  kgsLogoDataUrl?: string;
  screenshotBinary?: string;
  fileName?: string;
}

export interface IExportApplicationsBookArgs {
  items: IApplicationBookItem[];      // filtered applications mapped to export payload
  kgsLogoDataUrl?: string;
  fileName?: string;
}

export const exportApplicationPdf = async (args: IExportApplicationPdfArgs): Promise<void> => {

  const baseName: string =
    (args.fileName ?? args.application.Title)
      .replace(/\s+/g, "")   // remove ALL spaces
      .replace(/[\\/:*?"<>|]+/g, "-");  // sanitize illegal characters

  const safeFileName: string = `${baseName}_FactSheet.pdf`;

  const doc: React.ReactElement = (
    <Document>
      <ApplicationOnePager
        application={args.application}
        contract={args.contract}
        kgsLogoDataUrl={args.kgsLogoDataUrl}
        screenshotBinary={args.screenshotBinary}
      />
    </Document>
  );

  const blob: Blob = await pdf(doc).toBlob();
  saveAs(blob, safeFileName);

};


export const exportApplicationsBookPdf = async (
  args: IExportApplicationsBookArgs
): Promise<void> => {

  const baseName: string = (args.fileName ?? "Applications")
    .replace(/\s+/g, "")
    .replace(/[\\/:*?"<>|]+/g, "-");

  const safeFileName: string = `${baseName}_SummaryBook.pdf`;

  const doc: React.ReactElement = (
    <ApplicationBook
      items={args.items}
      kgsLogoDataUrl={args.kgsLogoDataUrl}
    />
  );

  const blob: Blob = await pdf(doc).toBlob();
  saveAs(blob, safeFileName);
};
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { IApplicationItem } from "../data/props";

// Helper: set a cell as a hyperlink with styled text
function setHyperlinkCell(cell: ExcelJS.Cell, text: string, href: string): void {
  cell.value = { text, hyperlink: href };
  cell.font = { color: { argb: "0563C1" }, underline: true };
}

export async function exportToExcel(apps: IApplicationItem[]): Promise<true> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Applications");

  worksheet.columns = [
    { header: "Application Title", key: "title", width: 30 },
    { header: "Description", key: "description", width: 100 },
    { header: "Primary POC", key: "primaryPoc", width: 30 },
    { header: "Stakeholders", key: "stakeholders", width: 50 },
    { header: "Managing Group", key: "managingGroup", width: 25 },
    { header: "URL", key: "appUrl", width: 18 },
    { header: "Contractor Name", key: "supportTeam", width: 30 },
    { header: "License Reqd", key: "licenseReqd", width: 20 },
    { header: "User Count", key: "userCount", width: 15 },
    { header: "Contract", key: "contract", width: 30 }
  ];

  apps.forEach((a) => {
    const row = worksheet.addRow({
      title: a.Title ?? "",
      description: a.description ?? "",
      primaryPoc: a.primaryPoc?.Title ?? "",
      stakeholders: a.stakeholders?.results?.map((bo) => bo.Title).join(", ") ?? "",
      managingGroup: a.managingGroup ?? "",
      appUrl: a.appUrl ?? "",
      supportTeam: a.supportTeam ?? "",
      licenseReqd: a.licenseReqd ?? "",
      userCount: typeof a.userCount === "number" ? a.userCount : 0,
      contract: a.contract?.Title ?? ""
    });

    if (a.appUrl) {
      setHyperlinkCell(row.getCell("appUrl"), "Open", a.appUrl);
    }

    row.getCell("description").alignment = { wrapText: true, vertical: "top" };
    row.getCell("stakeholders").alignment = { wrapText: true, vertical: "top" };
    row.getCell("contract").alignment = { wrapText: true, vertical: "top" };
  });

  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: "A1",
    to: "J1"
  };

  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "4472C4" }
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  const userCountCol = worksheet.getColumn("userCount");
  userCountCol.numFmt = "0";
  userCountCol.alignment = { horizontal: "center" };

  const today = new Date();
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, "");
  const fileName = `KGSApplications_${dateStr}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  saveAs(blob, fileName);
  return true;
}

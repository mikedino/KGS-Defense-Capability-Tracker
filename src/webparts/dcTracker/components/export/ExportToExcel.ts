import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ICapabilityItem } from "../common/props";

function setHyperlinkCell(cell: ExcelJS.Cell, text: string, href: string): void {
  cell.value = { text, hyperlink: href };
  cell.font = { color: { argb: "0563C1" }, underline: true };
}

export async function exportToExcel(capabilities: ICapabilityItem[]): Promise<true> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Capabilities");

  worksheet.columns = [
    { header: "Capability Title", key: "title", width: 30 },
    { header: "Capability Description", key: "description", width: 60 },
    { header: "Technical Capabilities", key: "capabilities", width: 60 },
    { header: "Link/URL", key: "link", width: 24 },
    { header: "Capability Status", key: "capStatus", width: 20 },
    { header: "Additional Notes", key: "notes", width: 50 },
    { header: "Platform", key: "platform", width: 20 },
    { header: "Hosting Environment", key: "hostingEnv", width: 24 },
    { header: "Connectivity", key: "connectivity", width: 20 },
    { header: "Compliance", key: "compliance", width: 20 },
    { header: "License Required?", key: "licenseReqd", width: 18 },
    { header: "Licensing Requirements", key: "licenseReqmts", width: 40 },
    { header: "APIs/Extensibility", key: "extensibility", width: 40 },
    { header: "Server Requirements", key: "serverReqmts", width: 28 },
    { header: "Coding Language", key: "codeLanguage", width: 22 },
    { header: "Backend", key: "backend", width: 22 },
    { header: "Contract", key: "contract", width: 30 }
  ];

  capabilities.forEach((cap) => {
    const row = worksheet.addRow({
      title: cap.Title ?? "",
      description: cap.description ?? "",
      capabilities: cap.capabilities ?? "",
      link: cap.link ?? "",
      capStatus: cap.capStatus ?? "",
      notes: cap.notes ?? "",
      platform: cap.platform ?? "",
      hostingEnv: cap.hostingEnv ?? "",
      connectivity: cap.connectivity ?? "",
      compliance: cap.compliance ?? "",
      licenseReqd: cap.licenseReqd ?? "",
      licenseReqmts: cap.licenseReqmts ?? "",
      extensibility: cap.extensibility ?? "",
      serverReqmts: cap.serverReqmts ?? "",
      codeLanguage: cap.codeLanguage ?? "",
      backend: cap.backend ?? "",
      contract: cap.contract?.Title ?? ""
    });

    if (cap.link) {
      setHyperlinkCell(row.getCell("link"), "Open", cap.link);
    }

    ["description", "capabilities", "notes", "licenseReqmts", "extensibility", "serverReqmts", "contract"].forEach((key) => {
      row.getCell(key).alignment = { wrapText: true, vertical: "top" };
    });
  });

  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: "A1",
    to: "Q1"
  };

  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "4472C4" }
    };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  const today = new Date();
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, "");
  const fileName = `KGSCapabilities_${dateStr}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  saveAs(blob, fileName);
  return true;
}

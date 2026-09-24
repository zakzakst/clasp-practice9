import type { FormItem, SheetInfos } from "./types/data";

const SHEET_ID = process.env.SHEET_ID!;

const getSheetData = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): Record<string, unknown>[] => {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map((header) => String(header ?? ""));
  return values.slice(1).map((row) => {
    const record: Record<string, unknown> = {};
    headers.forEach((header, index) => {
      record[header] = row[index] ?? "";
    });
    return record;
  });
};

const buildSpreadsheetResponse = (spreadsheetId: string): SheetInfos => {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  return spreadsheet.getSheets().map((sheet) => {
    const sheetName = sheet.getName();
    const formItems = getSheetData(sheet).map(
      (data) =>
        ({
          page: data["page"],
          label: data["label"],
          selector: data["selector"],
          type: data["type"],
          defaultValue: data["defaultValue"],
        }) as FormItem,
    );
    return {
      sheetName,
      formItems,
    };
  });
};

const doGet = (): GoogleAppsScript.Content.TextOutput => {
  if (!SHEET_ID) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "SHEET_ID is not configured." }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const response = buildSpreadsheetResponse(SHEET_ID);
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON,
  );
};

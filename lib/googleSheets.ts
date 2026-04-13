import { google } from "googleapis";

const requiredEnvVars = [
  "GOOGLE_CLIENT_EMAIL",
  "GOOGLE_PRIVATE_KEY",
  "GOOGLE_SHEETS_SPREADSHEET_ID",
] as const;

function getEnv(name: (typeof requiredEnvVars)[number]) {
  const value = process.env[name];

  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }

  return value;
}

function getOptionalEnv(name: string) {
  return process.env[name]?.trim();
}

function getSheetsClient() {
  const clientEmail = getEnv("GOOGLE_CLIENT_EMAIL");
  const privateKey = getEnv("GOOGLE_PRIVATE_KEY").replace(/\\n/g, "\n");

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: clientEmail,
      private_key: privateKey,
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

function escapeSheetName(sheetName: string) {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

export type SheetTab = {
  title: string;
  sheetId: number;
  rowCount: number;
  columnCount: number;
};

export type SheetWindow = {
  rows: string[][];
  hasNextPage: boolean;
  page: number;
  rowOffset: number;
};

export async function listSheetTabs(): Promise<SheetTab[]> {
  const sheets = getSheetsClient();
  const spreadsheetId = getEnv("GOOGLE_SHEETS_SPREADSHEET_ID");

  const response = await sheets.spreadsheets.get({
    spreadsheetId,
    fields:
      "sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))",
  });

  return (response.data.sheets ?? [])
    .map((sheet) => ({
      title: sheet.properties?.title ?? "Untitled Sheet",
      sheetId: sheet.properties?.sheetId ?? 0,
      rowCount: sheet.properties?.gridProperties?.rowCount ?? 0,
      columnCount: sheet.properties?.gridProperties?.columnCount ?? 0,
    }))
    .filter((sheet) => sheet.title);
}

export async function readSheetWindow(
  sheetName: string,
  page = 0,
  rowsPerPage = 10,
  columnsToShow = 8,
): Promise<SheetWindow> {
  const sheets = getSheetsClient();
  const spreadsheetId = getEnv("GOOGLE_SHEETS_SPREADSHEET_ID");
  const safePage = Math.max(0, page);
  const safeRowsPerPage = Math.max(1, rowsPerPage);
  const safeColumnsToShow = Math.max(1, Math.min(columnsToShow, 26));
  const startRow = safePage * safeRowsPerPage + 1;
  const endRow = startRow + safeRowsPerPage;
  const endColumn = String.fromCharCode(64 + safeColumnsToShow);
  const range = `${escapeSheetName(sheetName)}!A${startRow}:${endColumn}${endRow}`;

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
    majorDimension: "ROWS",
  });

  const values = (response.data.values ?? []).map((row) =>
    row.slice(0, safeColumnsToShow).map((cell) => String(cell ?? "").trim()),
  );

  return {
    rows: values.slice(0, safeRowsPerPage),
    hasNextPage: values.length > safeRowsPerPage,
    page: safePage,
    rowOffset: startRow,
  };
}

export async function appendLeadRow(row: string[]) {
  const sheets = getSheetsClient();
  const spreadsheetId = getEnv("GOOGLE_SHEETS_SPREADSHEET_ID");
  const sheetName = getOptionalEnv("GOOGLE_SHEETS_SHEET_NAME");

  if (!sheetName) {
    throw new Error("Missing required environment variable: GOOGLE_SHEETS_SHEET_NAME");
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${escapeSheetName(sheetName)}!A:G`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [row],
    },
  });
}

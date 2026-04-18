import { ImageResponse } from "next/og";
import { createElement } from "react";
import { NextResponse } from "next/server";
import { listSheetTabs, readSheetRange, readSheetWindow } from "@/lib/googleSheets";

type TelegramChat = {
  id: number;
};

type TelegramMessage = {
  message_id: number;
  chat: TelegramChat;
  text?: string;
  photo?: Array<Record<string, unknown>>;
};

type TelegramCallbackQuery = {
  id: string;
  data?: string;
  message?: TelegramMessage;
};

type TelegramUpdate = {
  message?: TelegramMessage;
  callback_query?: TelegramCallbackQuery;
};

type InlineKeyboardButton = {
  text: string;
  callback_data: string;
};

type DisplayRow = {
  rowLabel: string;
  content: string;
};

type RealTimeRow = {
  platform: string;
  dayShift: string;
  nightShift: string;
  midShift: string;
  total: string;
  isTotal?: boolean;
};

type RealTimePreview = {
  timestamp: string;
  headers: [string, string, string, string, string];
  rows: RealTimeRow[];
  rowOffset: number;
};

type MatrixSectionPreview = {
  title: string;
  subtitle: string;
  rows: string[][];
  headerRows: number;
  width: number;
  columnWidths?: number[];
  accent: string;
  badge: string;
};

const REAL_TIME_SECTION_COUNT = 5;

type ShiftingSummaryItem = {
  label: string;
  value: string;
};

type ShiftingEntry = {
  shift: string;
  platform: string;
  role: string;
  name: string;
  id: string;
  startDate: string;
  account: string;
};

type ShiftingSection = {
  title: string;
  dayEntries: ShiftingEntry[];
  midEntries: ShiftingEntry[];
  nightEntries: ShiftingEntry[];
};

type ShiftingShiftKind = "day" | "mid" | "night";

type ShiftingPreview = {
  summary: ShiftingSummaryItem[];
  sections: ShiftingSection[];
  platformIndex: number;
  currentSection: ShiftingSection;
  shiftKind: ShiftingShiftKind;
  entries: ShiftingEntry[];
  entryPage: number;
  totalEntryPages: number;
  pagedEntries: ShiftingEntry[];
};

type DailyTransactionMetric = {
  label: string;
  value: string;
};

type DailyTransactionEntry = {
  shift: string;
  name: string;
  platform: string;
  systemUser: string;
  metrics: DailyTransactionMetric[];
};

type DailyTransactionPlatform = {
  title: string;
  entries: DailyTransactionEntry[];
};

type DailyTransactionSummary = {
  totalPlatforms: number;
  totalStaff: number;
  dayCount: number;
  midCount: number;
  nightCount: number;
};

type DailyTransactionPreview = {
  summary: DailyTransactionSummary;
  headers: string[];
  platforms: DailyTransactionPlatform[];
  platformIndex: number;
  currentPlatform: DailyTransactionPlatform;
  shiftKind: ShiftingShiftKind;
  entries: DailyTransactionEntry[];
  entryPage: number;
  totalEntryPages: number;
  pagedEntries: DailyTransactionEntry[];
};

type BasicsStep = {
  title: string;
  subtitle: string;
  points: string[];
  accent: string;
  badge: string;
};

type BasicsPreview = {
  steps: BasicsStep[];
};

type WorkfolioEmailEntry = {
  id: string;
  name: string;
  email: string;
};

type WorkfolioEmailSection = {
  title: string;
  badge: string;
  entries: WorkfolioEmailEntry[];
};

type WorkfolioEmailPreview = {
  sections: WorkfolioEmailSection[];
  page: number;
  currentSection: WorkfolioEmailSection;
};

type SheetNavigation = {
  inline_keyboard: InlineKeyboardButton[][];
};

type DashboardKey = "withdraw" | "deposit";

const DASHBOARD_SHEET_TITLES = ["REAL TIME", "SHIFTING", "APRIL DAILY TRANSACTIONS"] as const;

function getDepositSpreadsheetId() {
  return process.env.GOOGLE_SHEETS_SPREADSHEET_ID_DEPOSIT?.trim() || "";
}

function getSpreadsheetIdForDashboard(dashboard: DashboardKey) {
  if (dashboard === "deposit") {
    const spreadsheetId = getDepositSpreadsheetId();

    if (!spreadsheetId) {
      throw new Error("Missing required environment variable: GOOGLE_SHEETS_SPREADSHEET_ID_DEPOSIT");
    }

    return spreadsheetId;
  }

  const spreadsheetId = process.env.GOOGLE_SHEETS_SPREADSHEET_ID?.trim();

  if (!spreadsheetId) {
    throw new Error("Missing required environment variable: GOOGLE_SHEETS_SPREADSHEET_ID");
  }

  return spreadsheetId;
}

function getAvailableDashboards(): DashboardKey[] {
  return getDepositSpreadsheetId() ? ["withdraw", "deposit"] : ["withdraw"];
}

function getDashboardLabel(dashboard: DashboardKey) {
  return dashboard === "deposit" ? "Deposit Dashboard" : "Withdraw Dashboard";
}

function getDashboardBadge(dashboard: DashboardKey) {
  return dashboard === "deposit" ? "🏦" : "💸";
}

function getTelegramBotToken() {
  const token = process.env.TELEGRAM_BOT_TOKEN?.trim();

  if (!token) {
    throw new Error("Missing required environment variable: TELEGRAM_BOT_TOKEN");
  }

  return token;
}

function getAllowedChatIds() {
  const raw = process.env.TELEGRAM_ALLOWED_CHAT_IDS?.trim();

  if (!raw) {
    return [];
  }

  return raw
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
}

function getPreviewRows() {
  return Math.max(1, Number(process.env.TELEGRAM_SHEET_PREVIEW_ROWS ?? "10"));
}

function getPreviewColumns() {
  return Math.max(1, Number(process.env.TELEGRAM_SHEET_PREVIEW_COLUMNS ?? "8"));
}

function isAllowedChat(chatId: number) {
  const allowedChatIds = getAllowedChatIds();

  if (allowedChatIds.length === 0) {
    return true;
  }

  return allowedChatIds.includes(String(chatId));
}

function cleanCell(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function shortenCell(value: string, maxLength = 18) {
  if (value.length <= maxLength) {
    return value;
  }

  return `${value.slice(0, maxLength - 1)}...`;
}

function normalizeHeader(value: string) {
  if (value.toUpperCase() === "PLATFROM") {
    return "PLATFORM";
  }

  return value;
}

function getSheetBadge(sheetTitle: string) {
  const title = sheetTitle.toLowerCase();

  if (title.includes("real time")) return "📊";
  if (title.includes("shift")) return "🔄";
  if (title.includes("account")) return "💳";
  if (title.includes("attendance")) return "🕒";
  if (title.includes("meeting")) return "📝";
  if (title.includes("email")) return "📧";
  if (title.includes("system")) return "🛠️";
  if (title.includes("workfolio")) return "🗂️";
  if (title.includes("april") || title.includes("march")) return "📅";
  if (title.includes("withdraw")) return "💸";

  return "📄";
}

function getSheetAccent(sheetTitle: string) {
  const palette = [
    "#2563eb",
    "#059669",
    "#7c3aed",
    "#ea580c",
    "#0f766e",
    "#dc2626",
    "#4338ca",
    "#0891b2",
  ];

  let hash = 0;

  for (const char of sheetTitle) {
    hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  }

  return palette[hash % palette.length];
}

function getSheetWindowConfig(sheetTitle: string) {
  const title = sheetTitle.toLowerCase();

  if (title.includes("account")) {
    return { rowsPerPage: 6, columnsToShow: 9 };
  }

  if (title.includes("attendance")) {
    return { rowsPerPage: 8, columnsToShow: 12 };
  }

  if (title.includes("daily transactions")) {
    return { rowsPerPage: 8, columnsToShow: 12 };
  }

  if (title.includes("april 1-15") || title.includes("april 16-30")) {
    return { rowsPerPage: 8, columnsToShow: 12 };
  }

  if (title.includes("workfolio&tl")) {
    return { rowsPerPage: 8, columnsToShow: 10 };
  }

  if (title.includes("meeting agenda")) {
    return { rowsPerPage: 8, columnsToShow: 6 };
  }

  return {
    rowsPerPage: getPreviewRows(),
    columnsToShow: getPreviewColumns(),
  };
}

async function withTimeout<T>(promise: Promise<T>, ms: number, label: string) {
  let timer: ReturnType<typeof setTimeout> | undefined;

  try {
    return await Promise.race([
      promise,
      new Promise<T>((_, reject) => {
        timer = setTimeout(() => reject(new Error(`${label} timed out.`)), ms);
      }),
    ]);
  } finally {
    if (timer) {
      clearTimeout(timer);
    }
  }
}

function buildDisplayRows(rows: string[][], rowOffset: number) {
  const visibleRows = rows
    .map((row, index) => ({
      rowLabel: String(rowOffset + index),
      content: row
        .map((cell) => cleanCell(cell || "-"))
        .filter(Boolean)
        .map((cell) => shortenCell(cell))
        .join("   |   "),
    }))
    .filter((row) => row.content.length > 0);

  if (visibleRows.length > 0) {
    return visibleRows;
  }

  return [
    {
      rowLabel: "-",
      content: "No data found in this range yet.",
    },
  ];
}

function isMatrixRowEmpty(row: string[]) {
  return row.every((cell) => cleanCell(cell ?? "").length === 0);
}

function trimMatrixRows(rows: string[][]) {
  const cleaned = rows.map((row) => row.map((cell) => cleanCell(cell ?? "")));
  const nonEmptyRows = cleaned.filter((row) => !isMatrixRowEmpty(row));

  if (nonEmptyRows.length === 0) {
    return [["No data available"]];
  }

  let maxColumnIndex = 0;

  for (const row of nonEmptyRows) {
    for (let index = row.length - 1; index >= 0; index -= 1) {
      if (row[index]) {
        maxColumnIndex = Math.max(maxColumnIndex, index);
        break;
      }
    }
  }

  return nonEmptyRows.map((row) =>
    Array.from({ length: maxColumnIndex + 1 }, (_, index) => row[index] ?? ""),
  );
}

async function getRealTimeSummaryPreview(): Promise<RealTimePreview> {
  const rows = await readSheetRange("REAL TIME", "A1:E26");
  const timestamp = cleanCell(rows[0]?.[0] ?? "REAL TIME");
  const rawHeaders = rows[1] ?? ["PLATFORM", "DAY SHIFT", "NIGHT SHIFT", "MID SHIFT", "TOTAL"];
  const headers = [
    normalizeHeader(rawHeaders[0] ?? "PLATFORM"),
    rawHeaders[1] ?? "DAY SHIFT",
    rawHeaders[2] ?? "NIGHT SHIFT",
    rawHeaders[3] ?? "MID SHIFT",
    rawHeaders[4] ?? "TOTAL",
  ] as [string, string, string, string, string];

  const dataRows = rows
    .slice(2)
    .filter((row) => cleanCell(row[0] ?? "").length > 0)
    .map((row) => ({
      platform: cleanCell(row[0] ?? "-"),
      dayShift: cleanCell(row[1] ?? "-"),
      nightShift: cleanCell(row[2] ?? "-"),
      midShift: cleanCell(row[3] ?? "-"),
      total: cleanCell(row[4] ?? "-"),
      isTotal: cleanCell(row[0] ?? "").toUpperCase() === "TOTAL",
    }));

  return {
    timestamp,
    headers,
    rows: dataRows,
    rowOffset: 3,
  };
}

async function getRealTimeMatrixSection(sectionIndex: number): Promise<MatrixSectionPreview> {
  switch (sectionIndex) {
    case 1: {
      const rows = trimMatrixRows(await readSheetRange("REAL TIME", "G1:V12"));
      return {
        title: "Daily Shift Codes",
        subtitle: "Section 2 of 5",
        rows,
        headerRows: 2,
        width: 1980,
        columnWidths: [150, ...Array.from({ length: Math.max(rows[0]?.length ?? 1, 2) - 1 }, () => 122)],
        accent: "#2563eb",
        badge: "📆",
      };
    }
    case 2: {
      const rows = trimMatrixRows(await readSheetRange("REAL TIME", "Y1:AC15"));
      return {
        title: "Team Leaders",
        subtitle: "Section 3 of 5",
        rows,
        headerRows: 1,
        width: 1800,
        columnWidths: [350, 160, 180, 170, 940],
        accent: "#7c3aed",
        badge: "👥",
      };
    }
    case 3: {
      const rows = trimMatrixRows(await readSheetRange("REAL TIME", "M35:Q50"));
      return {
        title: "Vietnam Staffs",
        subtitle: "Section 4 of 5",
        rows,
        headerRows: 2,
        width: 1480,
        columnWidths: [420, 210, 210, 210, 210],
        accent: "#059669",
        badge: "🇻🇳",
      };
    }
    case 4: {
      const rows = trimMatrixRows(await readSheetRange("REAL TIME", "A78:I120"));
      return {
        title: "Detailed Metrics",
        subtitle: "Section 5 of 5",
        rows,
        headerRows: 2,
        width: 1880,
        columnWidths: [500, 120, 120, 120, 150, 150, 150, 170, 120],
        accent: "#ea580c",
        badge: "📋",
      };
    }
    default:
      throw new Error(`Unsupported REAL TIME section: ${sectionIndex}`);
  }
}

function createShiftingEntry(row: string[], startIndex: number): ShiftingEntry | null {
  const shift = cleanCell(row[startIndex] ?? "");
  const platform = cleanCell(row[startIndex + 1] ?? "");
  const role = cleanCell(row[startIndex + 2] ?? "");
  const name = cleanCell(row[startIndex + 3] ?? "");
  const id = cleanCell(row[startIndex + 4] ?? "");
  const startDate = cleanCell(row[startIndex + 5] ?? "");
  const account = cleanCell(row[startIndex + 6] ?? "");

  if (!shift && !platform && !role && !name) {
    return null;
  }

  return {
    shift,
    platform,
    role,
    name,
    id,
    startDate,
    account,
  };
}

function isShiftingSectionHeader(row: string[]) {
  const e = cleanCell(row[4] ?? "");
  const f = cleanCell(row[5] ?? "");
  const g = cleanCell(row[6] ?? "");
  const h = cleanCell(row[7] ?? "");

  return Boolean(e) && !f && !g && !h;
}

async function getBasicsWithdrawPreview(): Promise<BasicsPreview> {
  const rows = await readSheetRange("BASICS WITHDARW", "A1:I8");
  const columns = [0, 3, 6];
  const accents = ["#2563eb", "#7c3aed", "#ea580c"];
  const badges = ["🔐", "🔔", "💸"];

  const steps = columns.map((startColumn, index) => ({
    title: cleanCell(rows[0]?.[startColumn] ?? `STEP-${index + 1}`),
    subtitle: cleanCell(rows[1]?.[startColumn] ?? ""),
    points: rows
      .slice(2)
      .map((row) => cleanCell(row[startColumn] ?? ""))
      .filter(Boolean),
    accent: accents[index],
    badge: badges[index],
  }));

  return { steps };
}

async function getWorkfolioEmailPreview(page: number): Promise<WorkfolioEmailPreview> {
  const rows = await readSheetRange("WORKFOLIO EMAIL", "A1:I120");
  const sectionDefs = [
    { title: cleanCell(rows[0]?.[0] ?? "DAY SHIFT"), badge: "🌤️", columns: [0, 1, 2] },
    { title: cleanCell(rows[0]?.[3] ?? "MID SHIFT"), badge: "🌇", columns: [3, 4, 5] },
    { title: cleanCell(rows[0]?.[6] ?? "NIGHT SHIFT"), badge: "🌙", columns: [6, 7, 8] },
  ];

  const sections = sectionDefs.map((section) => ({
    title: section.title,
    badge: section.badge,
    entries: rows
      .slice(2)
      .map((row) => ({
        id: cleanCell(row[section.columns[0]] ?? ""),
        name: cleanCell(row[section.columns[1]] ?? ""),
        email: cleanCell(row[section.columns[2]] ?? ""),
      }))
      .filter((entry) => entry.id || entry.name || entry.email),
  }));

  const safePage = Math.max(0, Math.min(page, sections.length - 1));

  return {
    sections,
    page: safePage,
    currentSection: sections[safePage],
  };
}

async function getShiftingData() {
  const rows = await readSheetRange("SHIFTING", "A1:R220");
  const summaryRows = rows.slice(1, 10);
  const summary = summaryRows
    .map((row) => ({
      label: cleanCell(row[0] ?? ""),
      value: cleanCell(row[1] ?? ""),
    }))
    .filter((item) => item.label && item.value);

  const sections: ShiftingSection[] = [];
  let currentSection: ShiftingSection | null = null;

  for (const row of rows.slice(1)) {
    if (isShiftingSectionHeader(row)) {
      const title = cleanCell(row[4] ?? "");

      if (title) {
        currentSection = {
          title,
          dayEntries: [],
          midEntries: [],
          nightEntries: [],
        };
        sections.push(currentSection);
      }

      continue;
    }

    if (!currentSection) {
      continue;
    }

    const dayEntry = createShiftingEntry(row, 4);
    const nightEntry = createShiftingEntry(row, 11);

    if (dayEntry && dayEntry.name) {
      const normalizedShift = dayEntry.shift.toLowerCase();

      if (normalizedShift.includes("中班") || normalizedShift.includes("mid")) {
        currentSection.midEntries.push(dayEntry);
      } else if (normalizedShift.includes("夜班") || normalizedShift.includes("night")) {
        currentSection.nightEntries.push(dayEntry);
      } else {
        currentSection.dayEntries.push(dayEntry);
      }
    }

    if (nightEntry && nightEntry.name) {
      const normalizedShift = nightEntry.shift.toLowerCase();

      if (normalizedShift.includes("中班") || normalizedShift.includes("mid")) {
        currentSection.midEntries.push(nightEntry);
      } else if (normalizedShift.includes("白班") || normalizedShift.includes("day")) {
        currentSection.dayEntries.push(nightEntry);
      } else {
        currentSection.nightEntries.push(nightEntry);
      }
    }
  }

  return {
    summary,
    sections,
  };
}

async function getShiftingPreview(
  platformIndex: number,
  shiftKind: ShiftingShiftKind,
  entryPage = 0,
): Promise<ShiftingPreview> {
  const { summary, sections } = await getShiftingData();
  const safePlatformIndex = Math.max(0, Math.min(platformIndex, Math.max(sections.length - 1, 0)));
  const fallbackSection: ShiftingSection = {
    title: "SHIFTING",
    dayEntries: [],
    midEntries: [],
    nightEntries: [],
  };
  const currentSection = sections[safePlatformIndex] ?? fallbackSection;
  const entries =
    shiftKind === "night"
      ? currentSection.nightEntries
      : shiftKind === "mid"
        ? currentSection.midEntries
        : currentSection.dayEntries;
  const entriesPerPage = 3;
  const totalEntryPages = Math.max(1, Math.ceil(entries.length / entriesPerPage));
  const safeEntryPage = Math.max(0, Math.min(entryPage, totalEntryPages - 1));
  const pagedEntries = entries.slice(
    safeEntryPage * entriesPerPage,
    safeEntryPage * entriesPerPage + entriesPerPage,
  );

  return {
    summary,
    sections,
    platformIndex: safePlatformIndex,
    currentSection,
    shiftKind,
    entries,
    entryPage: safeEntryPage,
    totalEntryPages,
    pagedEntries,
  };
}

function normalizeTransactionShift(shift: string) {
  const normalized = shift.toLowerCase();

  if (normalized.includes("中班") || normalized.includes("mid")) {
    return "mid";
  }

  if (normalized.includes("夜班") || normalized.includes("night")) {
    return "night";
  }

  return "day";
}

function getTransactionStats(metrics: DailyTransactionMetric[]) {
  const rdCount = metrics.filter((metric) => metric.value.toUpperCase() === "RD").length;
  const absCount = metrics.filter((metric) => metric.value.toUpperCase() === "ABS").length;
  const numericValues = metrics
    .map((metric) => Number(metric.value.replace(/,/g, "")))
    .filter((value) => Number.isFinite(value));
  const totalProcessed = numericValues.reduce((sum, value) => sum + value, 0);
  const peakValue = numericValues.length > 0 ? Math.max(...numericValues) : 0;
  const activeDays = metrics.filter((metric) => metric.value.length > 0).length;

  return {
    rdCount,
    absCount,
    numericValues,
    totalProcessed,
    peakValue,
    activeDays,
  };
}

async function getDailyTransactionData() {
  const rows = await readSheetRange("APRIL DAILY TRANSACTIONS", "A1:N260");
  const headerRow =
    rows.find((row) =>
      row.slice(4, 14).filter((cell) => cleanCell(cell ?? "").length > 0).length >= 6,
    ) ?? [];
  const headers = Array.from({ length: 10 }, (_, index) => cleanCell(headerRow[index + 4] ?? `${index + 1}`) || `${index + 1}`);
  const platformMap = new Map<string, DailyTransactionEntry[]>();
  let dayCount = 0;
  let midCount = 0;
  let nightCount = 0;

  for (const row of rows) {
    const shift = cleanCell(row[0] ?? "");
    const name = cleanCell(row[1] ?? "");
    const platform = cleanCell(row[2] ?? "");
    const systemUser = cleanCell(row[3] ?? "");

    if (!name || !platform) {
      continue;
    }

    if (
      name.toUpperCase() === "WD NAME" ||
      platform.toUpperCase() === "PLATFORM NAME" ||
      systemUser.toUpperCase() === "SYSTEM USER"
    ) {
      continue;
    }

    const metrics = headers.map((label, index) => ({
      label,
      value: cleanCell(row[index + 4] ?? ""),
    }));

    if (metrics.every((metric) => metric.value.length === 0) && !systemUser) {
      continue;
    }

    const entry: DailyTransactionEntry = {
      shift,
      name,
      platform,
      systemUser,
      metrics,
    };

    const normalizedShift = normalizeTransactionShift(shift);

    if (normalizedShift === "mid") {
      midCount += 1;
    } else if (normalizedShift === "night") {
      nightCount += 1;
    } else {
      dayCount += 1;
    }

    const existing = platformMap.get(platform) ?? [];
    existing.push(entry);
    platformMap.set(platform, existing);
  }

  const platforms = Array.from(platformMap.entries()).map(([title, entries]) => ({
    title,
    entries,
  }));

  return {
    headers,
    platforms,
    summary: {
      totalPlatforms: platforms.length,
      totalStaff: platforms.reduce((sum, platform) => sum + platform.entries.length, 0),
      dayCount,
      midCount,
      nightCount,
    },
  };
}

async function getDailyTransactionPreview(
  platformIndex: number,
  shiftKind: ShiftingShiftKind,
  entryPage = 0,
): Promise<DailyTransactionPreview> {
  const { headers, platforms, summary } = await getDailyTransactionData();
  const safePlatformIndex = Math.max(0, Math.min(platformIndex, Math.max(platforms.length - 1, 0)));
  const fallbackPlatform: DailyTransactionPlatform = { title: "APRIL DAILY TRANSACTIONS", entries: [] };
  const currentPlatform = platforms[safePlatformIndex] ?? fallbackPlatform;
  const entries = currentPlatform.entries.filter((entry) => normalizeTransactionShift(entry.shift) === shiftKind);
  const entriesPerPage = 3;
  const totalEntryPages = Math.max(1, Math.ceil(entries.length / entriesPerPage));
  const safeEntryPage = Math.max(0, Math.min(entryPage, totalEntryPages - 1));
  const pagedEntries = entries.slice(
    safeEntryPage * entriesPerPage,
    safeEntryPage * entriesPerPage + entriesPerPage,
  );

  return {
    summary,
    headers,
    platforms,
    platformIndex: safePlatformIndex,
    currentPlatform,
    shiftKind,
    entries,
    entryPage: safeEntryPage,
    totalEntryPages,
    pagedEntries,
  };
}

function buildSheetCaption(sheetTitle: string, _rowOffset: number, _rowCount: number, _rows: string[][]) {
  return `${getSheetBadge(sheetTitle)} ${sheetTitle}`;
}

function buildBasicsCaption() {
  return "📘 BASICS WITHDRAW\nStep by step guide\nRead-only training view";
}

function buildWorkfolioEmailCaption(preview: WorkfolioEmailPreview) {
  return `📧 WORKFOLIO EMAIL\n${preview.currentSection.badge} ${preview.currentSection.title}\nSection ${preview.page + 1} of ${preview.sections.length}`;
}

function buildRealTimeSummaryCaption(preview: RealTimePreview) {
  return `📊 REAL TIME
Live Operations Board
Updated ${preview.timestamp}`;
}

function buildRealTimeSectionCaption(preview: MatrixSectionPreview) {
  return `📊 REAL TIME
${preview.badge} ${preview.title}
Operational Detail View`;
}

function buildShiftingCaption(preview: ShiftingPreview) {
  const shiftLabel =
    preview.shiftKind === "night"
      ? "🌙 Night Shift"
      : preview.shiftKind === "mid"
        ? "🌇 Mid Shift"
        : "🌤️ Day Shift";
  const visibleCount = preview.pagedEntries.length;
  const pageLabel = preview.totalEntryPages > 1 ? `
Showing ${visibleCount} of ${preview.entries.length} · Page ${preview.entryPage + 1} of ${preview.totalEntryPages}` : "";

  return `🔄 SHIFTING COMMAND BOARD
${preview.currentSection.title}
${shiftLabel} · ${preview.entries.length} team member${preview.entries.length === 1 ? "" : "s"}${pageLabel}`;
}

function buildDailyTransactionCaption(preview: DailyTransactionPreview) {
  const shiftLabel =
    preview.shiftKind === "night"
      ? "🌙 Night Shift"
      : preview.shiftKind === "mid"
        ? "🌇 Mid Shift"
        : "🌤️ Day Shift";
  const totalProcessed = preview.entries
    .flatMap((entry) => entry.metrics)
    .map((metric) => Number(metric.value.replace(/,/g, "")))
    .filter((value) => Number.isFinite(value))
    .reduce((sum, value) => sum + value, 0);
  const pageLabel = preview.totalEntryPages > 1
    ? `\nShowing ${preview.pagedEntries.length} of ${preview.entries.length} · Page ${preview.entryPage + 1} of ${preview.totalEntryPages}`
    : "";

  return `💹 DAILY TRANSACTIONS COMMAND BOARD
${preview.currentPlatform.title}
${shiftLabel} · ${preview.entries.length} staff · Total Processed ${totalProcessed.toLocaleString("en-US")}${pageLabel}`;
}

function buildSheetNavigation(
  dashboard: DashboardKey,
  sheetIndex: number,
  page: number,
  totalPages: number,
  rowLabel?: string,
): SheetNavigation {
  const inlineKeyboard: InlineKeyboardButton[][] = [];
  const navigationButtons: InlineKeyboardButton[] = [];
  const safeTotalPages = Math.max(1, totalPages);
  const lastPage = safeTotalPages - 1;

  if (page > 0) {
    navigationButtons.push({
      text: "⏮️ First",
      callback_data: `sheet:${dashboard}:${sheetIndex}:0`,
    });
    navigationButtons.push({
      text: "⬅️ Previous",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${page - 1}`,
    });
  }

  if (page < lastPage) {
    navigationButtons.push({
      text: "Next ➡️",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${page + 1}`,
    });
    navigationButtons.push({
      text: "Last ⏭️",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${lastPage}`,
    });
  }

  if (navigationButtons.length > 0) {
    inlineKeyboard.push(navigationButtons);
  }

  const jumpButtons: InlineKeyboardButton[] = [];

  if (page >= 10) {
    jumpButtons.push({
      text: "⏪ -10",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${Math.max(0, page - 10)}`,
    });
  }

  if (rowLabel) {
    jumpButtons.push({
      text: `🔢 ${rowLabel}`,
      callback_data: "noop:0",
    });
  }

  if (page + 10 < safeTotalPages) {
    jumpButtons.push({
      text: "+10 ⏩",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${Math.min(lastPage, page + 10)}`,
    });
  }

  if (jumpButtons.length > 0) {
    inlineKeyboard.push(jumpButtons);
  }

  inlineKeyboard.push([
    { text: `↩ ${getDashboardLabel(dashboard)}`, callback_data: `menu:${dashboard}` },
    { text: "🧭 Workspaces", callback_data: "home:0" },
  ]);

  return { inline_keyboard: inlineKeyboard };
}

function buildSectionNavigation(
  dashboard: DashboardKey,
  sheetIndex: number,
  page: number,
  totalSections: number,
): SheetNavigation {
  const inlineKeyboard: InlineKeyboardButton[][] = [];
  const navigationButtons: InlineKeyboardButton[] = [];

  if (page > 0) {
    navigationButtons.push({
      text: "⬅️ Previous",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${page - 1}`,
    });
  }

  if (page + 1 < totalSections) {
    navigationButtons.push({
      text: "Next ➡️",
      callback_data: `sheet:${dashboard}:${sheetIndex}:${page + 1}`,
    });
  }

  if (navigationButtons.length > 0) {
    inlineKeyboard.push(navigationButtons);
  }

  inlineKeyboard.push([
    { text: `↩ ${getDashboardLabel(dashboard)}`, callback_data: `menu:${dashboard}` },
    { text: "🧭 Workspaces", callback_data: "home:0" },
  ]);

  return { inline_keyboard: inlineKeyboard };
}

function buildShiftingPlatformKeyboard(sheetIndex: number, sections: ShiftingSection[]): SheetNavigation {
  const rows: InlineKeyboardButton[][] = [];

  for (let index = 0; index < sections.length; index += 2) {
    rows.push(
      sections.slice(index, index + 2).map((section, offset) => ({
        text: `🏷️ ${shortenCell(section.title, 18)}`,
        callback_data: `shiftplatform:${sheetIndex}:${index + offset}`,
      })),
    );
  }

  rows.push([
    { text: "↩ Withdraw Dashboard", callback_data: "menu:withdraw" },
    { text: "🧭 Workspaces", callback_data: "home:0" },
  ]);

  return { inline_keyboard: rows };
}

function buildShiftingShiftKeyboard(
  sheetIndex: number,
  platformIndex: number,
  section: ShiftingSection,
  shiftKind: ShiftingShiftKind,
  totalPlatforms: number,
  entryPage = 0,
  totalEntryPages = 1,
): SheetNavigation {
  const rows: InlineKeyboardButton[][] = [];
  const platformNav: InlineKeyboardButton[] = [];

  if (platformIndex > 0) {
    platformNav.push({
      text: "◀ Previous",
      callback_data: `shiftview:${sheetIndex}:${platformIndex - 1}:${shiftKind}:0`,
    });
  }

  if (platformIndex + 1 < totalPlatforms) {
    platformNav.push({
      text: "Next ▶",
      callback_data: `shiftview:${sheetIndex}:${platformIndex + 1}:${shiftKind}:0`,
    });
  }

  if (platformNav.length > 0) {
    rows.push(platformNav);
  }

  if (totalEntryPages > 1) {
    const pageNav: InlineKeyboardButton[] = [];

    if (entryPage > 0) {
      pageNav.push({
        text: "◀ Previous Page",
        callback_data: `shiftview:${sheetIndex}:${platformIndex}:${shiftKind}:${entryPage - 1}`,
      });
    }

    if (entryPage + 1 < totalEntryPages) {
      pageNav.push({
        text: "Next Page ▶",
        callback_data: `shiftview:${sheetIndex}:${platformIndex}:${shiftKind}:${entryPage + 1}`,
      });
    }

    if (pageNav.length > 0) {
      rows.push(pageNav);
    }
  }

  rows.push([
    {
      text: `🌤️ Day (${section.dayEntries.length})`,
      callback_data: `shiftview:${sheetIndex}:${platformIndex}:day:0`,
    },
    {
      text: `🌇 Mid (${section.midEntries.length})`,
      callback_data: `shiftview:${sheetIndex}:${platformIndex}:mid:0`,
    },
    {
      text: `🌙 Night (${section.nightEntries.length})`,
      callback_data: `shiftview:${sheetIndex}:${platformIndex}:night:0`,
    },
  ]);

  rows.push([
    { text: "🗂 Platform List", callback_data: `shiftplatforms:${sheetIndex}` },
    { text: "↩ Withdraw Dashboard", callback_data: "menu:withdraw" },
  ]);

  return { inline_keyboard: rows };
}

function buildDailyTransactionPlatformKeyboard(
  sheetIndex: number,
  platforms: DailyTransactionPlatform[],
): SheetNavigation {
  const rows: InlineKeyboardButton[][] = [];

  for (let index = 0; index < platforms.length; index += 2) {
    rows.push(
      platforms.slice(index, index + 2).map((platform, offset) => ({
        text: `🏷️ ${shortenCell(platform.title, 18)}`,
        callback_data: `txplatform:${sheetIndex}:${index + offset}:0`,
      })),
    );
  }

  rows.push([
    { text: "↩ Withdraw Dashboard", callback_data: "menu:withdraw" },
    { text: "🧭 Workspaces", callback_data: "home:0" },
  ]);

  return { inline_keyboard: rows };
}

function buildDailyTransactionShiftKeyboard(
  sheetIndex: number,
  platformIndex: number,
  platforms: DailyTransactionPlatform[],
  shiftKind: ShiftingShiftKind,
  entryPage = 0,
): SheetNavigation {
  const rows: InlineKeyboardButton[][] = [];
  const currentPlatform = platforms[platformIndex];

  if (!currentPlatform) {
    return { inline_keyboard: [[{ text: "↩ Withdraw Dashboard", callback_data: "menu:withdraw" }]] };
  }

  const platformNav: InlineKeyboardButton[] = [];

  if (platformIndex > 0) {
    platformNav.push({
      text: "◀ Previous",
      callback_data: `txview:${sheetIndex}:${platformIndex - 1}:${shiftKind}:0`,
    });
  }

  if (platformIndex + 1 < platforms.length) {
    platformNav.push({
      text: "Next ▶",
      callback_data: `txview:${sheetIndex}:${platformIndex + 1}:${shiftKind}:0`,
    });
  }

  if (platformNav.length > 0) {
    rows.push(platformNav);
  }

  const filteredEntries = currentPlatform.entries.filter((entry) => normalizeTransactionShift(entry.shift) === shiftKind);
  const perPage = 3;
  const totalPages = Math.max(1, Math.ceil(filteredEntries.length / perPage));
  const safePage = Math.max(0, Math.min(entryPage, totalPages - 1));

  if (totalPages > 1) {
    const pageNav: InlineKeyboardButton[] = [];

    if (safePage > 0) {
      pageNav.push({
        text: "◀ Previous Page",
        callback_data: `txview:${sheetIndex}:${platformIndex}:${shiftKind}:${safePage - 1}`,
      });
    }

    if (safePage + 1 < totalPages) {
      pageNav.push({
        text: "Next Page ▶",
        callback_data: `txview:${sheetIndex}:${platformIndex}:${shiftKind}:${safePage + 1}`,
      });
    }

    if (pageNav.length > 0) {
      rows.push(pageNav);
    }
  }

  rows.push([
    {
      text: `🌤️ Day (${currentPlatform.entries.filter((entry) => normalizeTransactionShift(entry.shift) === "day").length})`,
      callback_data: `txview:${sheetIndex}:${platformIndex}:day:0`,
    },
    {
      text: `🌇 Mid (${currentPlatform.entries.filter((entry) => normalizeTransactionShift(entry.shift) === "mid").length})`,
      callback_data: `txview:${sheetIndex}:${platformIndex}:mid:0`,
    },
    {
      text: `🌙 Night (${currentPlatform.entries.filter((entry) => normalizeTransactionShift(entry.shift) === "night").length})`,
      callback_data: `txview:${sheetIndex}:${platformIndex}:night:0`,
    },
  ]);

  rows.push([
    { text: "🗂 Platform List", callback_data: `txplatforms:${sheetIndex}` },
    { text: "↩ Withdraw Dashboard", callback_data: "menu:withdraw" },
  ]);

  return { inline_keyboard: rows };
}

async function sendKeyboardMessage(chatId: number, text: string, replyMarkup: SheetNavigation) {
  await callTelegram("sendMessage", {
    chat_id: chatId,
    text,
    reply_markup: replyMarkup,
  });
}

async function showShiftingPlatformMenu(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "Missing message context.");
    return;
  }

  const { summary, sections } = await getShiftingData();
  await answerCallbackQuery(callbackQuery.id, "Loading...");
  const loadingMessageId = await sendStatusMessage(message.chat.id, "⏳ Loading SHIFTING overview...");

  try {
    const imageBuffer = await renderShiftingOverviewImage(summary, sections.length);
    await deleteTelegramMessage(message.chat.id, message.message_id).catch(() => undefined);
    await sendTelegramPhoto(
      message.chat.id,
      imageBuffer,
      `🔄 SHIFTING COMMAND CENTER
All platforms overview
${sections.length} platform${sections.length === 1 ? "" : "s"} ready`,
      buildShiftingPlatformKeyboard(sheetIndex, sections),
    );
  } finally {
    if (loadingMessageId) {
      await deleteTelegramMessage(message.chat.id, loadingMessageId).catch(() => undefined);
    }
  }
}

async function showShiftingShiftMenu(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
  platformIndex: number,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "Missing message context.");
    return;
  }

  const { sections } = await getShiftingData();
  const safePlatformIndex = Math.max(0, Math.min(platformIndex, Math.max(sections.length - 1, 0)));
  const currentSection = sections[safePlatformIndex];

  if (!currentSection) {
    await answerCallbackQuery(callbackQuery.id, "Platform not found.");
    return;
  }

  await answerCallbackQuery(callbackQuery.id);
  await deleteTelegramMessage(message.chat.id, message.message_id).catch(() => undefined);

  await sendKeyboardMessage(
    message.chat.id,
    `🔄 SHIFTING
${currentSection.title}
Choose Day, Mid, or Night.`,
    buildShiftingShiftKeyboard(
      sheetIndex,
      safePlatformIndex,
      currentSection,
      currentSection.dayEntries.length > 0 ? "day" : currentSection.midEntries.length > 0 ? "mid" : "night",
      sections.length,
      0,
      1,
    ),
  );
}

async function showShiftingView(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
  platformIndex: number,
  shiftKind: ShiftingShiftKind,
  entryPage = 0,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "The original message is no longer available.");
    return;
  }

  const preview = await getShiftingPreview(platformIndex, shiftKind, entryPage);
  await answerCallbackQuery(callbackQuery.id, "Loading...");
  const loadingMessageId = await sendStatusMessage(
    message.chat.id,
    `⏳ Loading ${preview.currentSection.title} ${shiftKind === "mid" ? "mid" : shiftKind} shift...`,
  );

  try {
    const imageBuffer = await renderShiftingImage(preview);
    const caption = buildShiftingCaption(preview);
    const replyMarkup = buildShiftingShiftKeyboard(
      sheetIndex,
      preview.platformIndex,
      preview.currentSection,
      shiftKind,
      preview.sections.length,
      preview.entryPage,
      preview.totalEntryPages,
    );

    await deleteTelegramMessage(message.chat.id, message.message_id).catch(() => undefined);
    await sendTelegramPhoto(message.chat.id, imageBuffer, caption, replyMarkup);
  } finally {
    if (loadingMessageId) {
      await deleteTelegramMessage(message.chat.id, loadingMessageId).catch(() => undefined);
    }
  }
}

async function showDailyTransactionPlatformMenu(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "Missing message context.");
    return;
  }

  const { summary, platforms } = await getDailyTransactionData();
  await answerCallbackQuery(callbackQuery.id, "Loading...");
  const loadingMessageId = await sendStatusMessage(
    message.chat.id,
    "⏳ Loading daily transactions overview...",
  );

  try {
    const imageBuffer = await renderDailyTransactionOverviewImage(summary);
    await deleteTelegramMessage(message.chat.id, message.message_id).catch(() => undefined);
    await sendTelegramPhoto(
      message.chat.id,
      imageBuffer,
      `💹 DAILY TRANSACTIONS COMMAND CENTER
Choose a platform to continue
Then open Day, Mid, or Night
${platforms.length} platform${platforms.length === 1 ? "" : "s"} ready`,
      buildDailyTransactionPlatformKeyboard(sheetIndex, platforms),
    );
  } finally {
    if (loadingMessageId) {
      await deleteTelegramMessage(message.chat.id, loadingMessageId).catch(() => undefined);
    }
  }
}

async function showDailyTransactionShiftMenu(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
  platformIndex: number,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "Missing message context.");
    return;
  }

  const { platforms } = await getDailyTransactionData();
  const safePlatformIndex = Math.max(0, Math.min(platformIndex, Math.max(platforms.length - 1, 0)));
  const currentPlatform = platforms[safePlatformIndex];

  if (!currentPlatform) {
    await answerCallbackQuery(callbackQuery.id, "Platform not found.");
    return;
  }

  await answerCallbackQuery(callbackQuery.id);
  await deleteTelegramMessage(message.chat.id, message.message_id).catch(() => undefined);

  const defaultShift: ShiftingShiftKind = currentPlatform.entries.some((entry) => normalizeTransactionShift(entry.shift) === "day")
    ? "day"
    : currentPlatform.entries.some((entry) => normalizeTransactionShift(entry.shift) === "mid")
      ? "mid"
      : "night";

  await sendKeyboardMessage(
    message.chat.id,
    `💹 DAILY TRANSACTIONS
${currentPlatform.title}
Choose Day, Mid, or Night to view all staff transactions.`,
    buildDailyTransactionShiftKeyboard(
      sheetIndex,
      safePlatformIndex,
      platforms,
      defaultShift,
      0,
    ),
  );
}

async function showDailyTransactionView(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
  platformIndex: number,
  shiftKind: ShiftingShiftKind,
  entryPage = 0,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "The original message is no longer available.");
    return;
  }

  const preview = await getDailyTransactionPreview(platformIndex, shiftKind, entryPage);
  await answerCallbackQuery(callbackQuery.id, "Loading...");
  const loadingMessageId = await sendStatusMessage(
    message.chat.id,
    `⏳ Loading ${preview.currentPlatform.title} ${shiftKind} transactions...`,
  );

  try {
    const imageBuffer = await renderDailyTransactionEntryImage(preview);
    await deleteTelegramMessage(message.chat.id, message.message_id).catch(() => undefined);
    await sendTelegramPhoto(
      message.chat.id,
      imageBuffer,
      buildDailyTransactionCaption(preview),
      buildDailyTransactionShiftKeyboard(
        sheetIndex,
        preview.platformIndex,
        preview.platforms,
        preview.shiftKind,
        preview.entryPage,
      ),
    );
  } finally {
    if (loadingMessageId) {
      await deleteTelegramMessage(message.chat.id, loadingMessageId).catch(() => undefined);
    }
  }
}

async function callTelegram(method: string, payload: Record<string, unknown>) {
  const token = getTelegramBotToken();
  const response = await fetch(`https://api.telegram.org/bot${token}/${method}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  const body = (await response.json()) as { ok?: boolean; result?: unknown; description?: string };

  if (!response.ok) {
    throw new Error(
      `Telegram API ${method} failed: ${response.status} ${body.description ?? "Unknown error"}`,
    );
  }

  if (body.ok === false) {
    throw new Error(`Telegram API ${method} failed: ${body.description ?? "Unknown error"}`);
  }

  return body.result;
}

async function sendTelegramPhoto(
  chatId: number,
  imageBuffer: ArrayBuffer,
  caption: string,
  replyMarkup: SheetNavigation,
) {
  const token = getTelegramBotToken();
  const formData = new FormData();

  formData.append("chat_id", String(chatId));
  formData.append(
    "photo",
    new Blob([imageBuffer], { type: "image/png" }),
    "sheet-preview.png",
  );
  formData.append("caption", caption);
  formData.append("reply_markup", JSON.stringify(replyMarkup));

  const response = await fetch(`https://api.telegram.org/bot${token}/sendPhoto`, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Telegram API sendPhoto failed: ${response.status} ${body}`);
  }
}

async function answerCallbackQuery(callbackQueryId: string, text?: string) {
  await callTelegram("answerCallbackQuery", {
    callback_query_id: callbackQueryId,
    text,
  });
}

async function deleteTelegramMessage(chatId: number, messageId: number) {
  await callTelegram("deleteMessage", {
    chat_id: chatId,
    message_id: messageId,
  });
}

async function sendStatusMessage(chatId: number, text: string) {
  const result = (await callTelegram("sendMessage", {
    chat_id: chatId,
    text,
    disable_notification: true,
  })) as { message_id?: number } | undefined;

  return result?.message_id;
}

function getLoadingMessage(sheetTitle: string) {
  switch (sheetTitle) {
    case "REAL TIME":
      return "⏳ Loading REAL TIME dashboard...";
    case "SHIFTING":
      return "⏳ Loading SHIFTING dashboard...";
    case "APRIL DAILY TRANSACTIONS":
      return "⏳ Loading daily transactions board...";
    case "MAIN ALL ACCOUNTS":
    case "MAIN USED ACCOUNTS":
    case "SYSTEM ACCOUNTS":
      return "⏳ Loading account dashboard...";
    case "WORKFOLIO EMAIL":
      return "⏳ Loading workfolio email directory...";
    case "BASICS WITHDARW":
      return "⏳ Loading basics guide...";
    default:
      return `⏳ Loading ${sheetTitle}...`;
  }
}

async function getDashboardSheets(dashboard: DashboardKey) {
  const sheets = await listSheetTabs(getSpreadsheetIdForDashboard(dashboard));

  if (dashboard === "deposit") {
    return sheets;
  }

  return DASHBOARD_SHEET_TITLES.map((title) => sheets.find((sheet) => sheet.title === title)).filter(
    (sheet): sheet is NonNullable<(typeof sheets)[number]> => Boolean(sheet),
  );
}

async function buildSheetKeyboard(dashboard: DashboardKey) {
  const sheets = await getDashboardSheets(dashboard);
  const rows: InlineKeyboardButton[][] = sheets.map((sheet, index) => [{
    text:
      dashboard === "withdraw"
        ? sheet.title === "REAL TIME"
          ? "📊 Real-Time Command Center"
          : sheet.title === "SHIFTING"
            ? "🔄 Shifting Command Center"
            : "💹 Daily Transactions Center"
        : `${getSheetBadge(sheet.title)} ${shortenCell(sheet.title, 28)}`,
    callback_data: `sheet:${dashboard}:${index}:0`,
  }]);

  rows.push([{ text: "🧭 Switch Workspace", callback_data: "home:0" }]);

  return {
    inline_keyboard: rows,
  };
}

async function buildHomeKeyboard() {
  return {
    inline_keyboard: getAvailableDashboards().map((dashboard) => [{
      text: `${getDashboardBadge(dashboard)} ${getDashboardLabel(dashboard)}`,
      callback_data: `menu:${dashboard}`,
    }]),
  };
}

async function renderGenericSheetImage(
  sheetTitle: string,
  rowOffset: number,
  rowCount: number,
  rows: string[][],
) {
  const cleanedRows = rows.map((row) => row.map((cell) => cleanCell(cell || "")));
  const visibleRows = cleanedRows.filter((row) => row.some((cell) => cell.length > 0));
  const tableRows = visibleRows.length > 0 ? visibleRows : [["No data available"]];
  const columnCount = Math.max(...tableRows.map((row) => row.length), 1);
  const accent = getSheetAccent(sheetTitle);
  const badge = getSheetBadge(sheetTitle);
  const width = Math.min(1850, Math.max(1280, 320 + columnCount * 210));
  const headerRows = rowOffset === 1 ? 1 : 0;
  const bodyRowCount = Math.max(tableRows.length - headerRows, 0);
  const height = Math.max(920, 250 + headerRows * 82 + bodyRowCount * 74);
  const normalizedRows = tableRows.map((row) =>
    Array.from({ length: columnCount }, (_, index) => row[index] ?? ""),
  );

  const measuredWidths = Array.from({ length: columnCount }, (_, columnIndex) => {
    const longest = Math.max(
      10,
      ...normalizedRows.map((row) => Math.min((row[columnIndex] ?? "").length, 26)),
    );

    return longest;
  });
  const totalUnits = measuredWidths.reduce((sum, value) => sum + value, 0);
  const usableWidth = width - 56;
  const columnWidths = measuredWidths.map((value, index) => {
    const computed = Math.floor((value / totalUnits) * usableWidth);
    const minWidth = index === 0 ? 180 : 120;
    return Math.max(minWidth, computed);
  });

  const adjustedColumnWidths = columnWidths.map((value, index) => {
    if (index !== columnWidths.length - 1) {
      return value;
    }

    const consumed = columnWidths.slice(0, -1).reduce((sum, widthValue) => sum + widthValue, 0);
    return usableWidth - consumed;
  });

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #f8fafc 0%, #eff6ff 100%)",
          color: "#111827",
          fontFamily: "Georgia, serif",
          boxSizing: "border-box",
          padding: "28px",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: `linear-gradient(135deg, ${accent} 0%, #1f2937 100%)`,
              color: "#ffffff",
              padding: "22px 28px",
              borderRadius: "28px 28px 0 0",
              border: `3px solid ${accent}`,
              borderBottom: "0",
            },
          },
          [
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "46px",
                  fontWeight: 700,
                },
              },
              `${badge} ${sheetTitle}`,
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "22px",
                  marginTop: "8px",
                },
              },
              `Rows ${rowOffset}-${rows.length > 0 ? rowOffset + rows.length - 1 : rowOffset} of ${rowCount}`,
            ),
          ],
        ),
        ...normalizedRows.map((row, rowIndex) =>
          createElement(
            "div",
            {
              key: `row-${rowIndex}`,
              style: {
                display: "flex",
                background:
                  rowIndex < headerRows
                    ? accent
                    : rowIndex % 2 === 0
                      ? "#ffffff"
                      : "#f7f8fc",
                color: rowIndex < headerRows ? "#ffffff" : "#111827",
                borderLeft: `3px solid ${accent}`,
                borderRight: `3px solid ${accent}`,
                borderBottom: "2px solid #cbd5e1",
              },
            },
            row.map((cell, columnIndex) =>
              createElement(
                "div",
                {
                  key: `cell-${rowIndex}-${columnIndex}`,
                  style: {
                    width: `${adjustedColumnWidths[columnIndex]}px`,
                    padding: rowIndex < headerRows ? "14px 10px" : "12px 10px",
                    borderRight: columnIndex === columnCount - 1 ? "0" : "2px solid #cbd5e1",
                    textAlign: columnIndex === 0 ? "left" : "center",
                    fontSize: rowIndex < headerRows ? "24px" : "22px",
                    fontWeight: rowIndex < headerRows ? 700 : columnIndex === 0 ? 700 : 600,
                    whiteSpace: "pre-wrap",
                    wordBreak: "break-word",
                    lineHeight: 1.15,
                    minHeight: rowIndex < headerRows ? "64px" : "60px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: columnIndex === 0 ? "flex-start" : "center",
                  },
                },
                cell || (rowIndex < headerRows ? `Col ${columnIndex + 1}` : "-"),
              ),
            ),
          ),
        ),
      ],
    ),
    {
      width,
      height,
    },
  );

  return image.arrayBuffer();
}

async function renderBasicsWithdrawImage(preview: BasicsPreview) {
  const width = 1800;
  const height = 1200;

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%)",
          color: "#111827",
          fontFamily: "Georgia, serif",
          boxSizing: "border-box",
          padding: "30px",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: "linear-gradient(135deg, #1d4ed8 0%, #1f2937 100%)",
              color: "#ffffff",
              padding: "24px 28px",
              borderRadius: "30px",
              marginBottom: "24px",
            },
          },
          [
            createElement("div", { key: "title", style: { fontSize: "52px", fontWeight: 700 } }, "📘 BASICS WITHDRAW"),
            createElement("div", { key: "sub", style: { fontSize: "24px", marginTop: "8px" } }, "Quick guide for new staff and TLs"),
          ],
        ),
        createElement(
          "div",
          {
            key: "cards",
            style: {
              display: "flex",
              gap: "22px",
              flex: 1,
            },
          },
          preview.steps.map((step) =>
            createElement(
              "div",
              {
                key: step.title,
                style: {
                  display: "flex",
                  flexDirection: "column",
                  width: "33%",
                  background: "#ffffff",
                  borderRadius: "26px",
                  overflow: "hidden",
                  border: `3px solid ${step.accent}`,
                  boxShadow: "0 12px 30px rgba(15, 23, 42, 0.10)",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "card-header",
                    style: {
                      background: `linear-gradient(135deg, ${step.accent} 0%, #1f2937 100%)`,
                      color: "#ffffff",
                      padding: "20px 22px",
                    },
                  },
                  [
                    createElement("div", { key: "title", style: { fontSize: "30px", fontWeight: 700 } }, `${step.badge} ${step.title}`),
                    createElement("div", { key: "subtitle", style: { fontSize: "18px", marginTop: "8px" } }, step.subtitle),
                  ],
                ),
                createElement(
                  "div",
                  {
                    key: "points",
                    style: {
                      display: "flex",
                      flexDirection: "column",
                      padding: "18px",
                      gap: "14px",
                    },
                  },
                  step.points.map((point, index) =>
                    createElement(
                      "div",
                      {
                        key: `${step.title}-${index}`,
                        style: {
                          display: "flex",
                          background: index % 2 === 0 ? "#f8fafc" : "#eef2ff",
                          borderRadius: "18px",
                          padding: "14px 16px",
                          fontSize: "18px",
                          lineHeight: 1.25,
                        },
                      },
                      point,
                    ),
                  ),
                ),
              ],
            ),
          ),
        ),
      ],
    ),
    { width, height },
  );

  return image.arrayBuffer();
}

async function renderWorkfolioEmailImage(preview: WorkfolioEmailPreview) {
  const section = preview.currentSection;
  const width = 1500;
  const height = Math.max(980, 260 + section.entries.length * 72);
  const accent = preview.page === 0 ? "#2563eb" : preview.page === 1 ? "#7c3aed" : "#0f766e";

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #f8fafc 0%, #eff6ff 100%)",
          color: "#111827",
          fontFamily: "Georgia, serif",
          boxSizing: "border-box",
          padding: "28px",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: `linear-gradient(135deg, ${accent} 0%, #1f2937 100%)`,
              color: "#ffffff",
              padding: "22px 28px",
              borderRadius: "28px 28px 0 0",
              border: `3px solid ${accent}`,
              borderBottom: "0",
            },
          },
          [
            createElement("div", { key: "title", style: { fontSize: "46px", fontWeight: 700 } }, `📧 ${section.badge} ${section.title}`),
            createElement("div", { key: "sub", style: { fontSize: "22px", marginTop: "8px" } }, `Email contacts · Section ${preview.page + 1} of ${preview.sections.length}`),
          ],
        ),
        createElement(
          "div",
          {
            key: "thead",
            style: {
              display: "flex",
              background: accent,
              color: "#ffffff",
              borderLeft: `3px solid ${accent}`,
              borderRight: `3px solid ${accent}`,
              borderBottom: "3px solid #1f2937",
            },
          },
          [
            ["ID", 220],
            ["NAME", 540],
            ["EMAIL ADDRESS", 680],
          ].map(([label, widthValue], index) =>
            createElement(
              "div",
              {
                key: String(label),
                style: {
                  width: `${widthValue}px`,
                  padding: "14px 12px",
                  borderRight: index === 2 ? "0" : "2px solid #1f2937",
                  textAlign: index === 0 ? "center" : "left",
                  fontSize: "24px",
                  fontWeight: 700,
                },
              },
              String(label),
            ),
          ),
        ),
        ...section.entries.map((entry, index) =>
          createElement(
            "div",
            {
              key: `${entry.id}-${index}`,
              style: {
                display: "flex",
                background: index % 2 === 0 ? "#ffffff" : "#f7f8fc",
                borderLeft: `3px solid ${accent}`,
                borderRight: `3px solid ${accent}`,
                borderBottom: "2px solid #cbd5e1",
              },
            },
            [
              [entry.id || "-", 220, "center"],
              [entry.name || "-", 540, "left"],
              [entry.email || "-", 680, "left"],
            ].map(([value, widthValue, align], cellIndex) =>
              createElement(
                "div",
                {
                  key: `${entry.id}-${cellIndex}`,
                  style: {
                    width: `${widthValue}px`,
                    padding: "14px 12px",
                    borderRight: cellIndex === 2 ? "0" : "2px solid #cbd5e1",
                    textAlign: String(align),
                    fontSize: "22px",
                    fontWeight: cellIndex === 1 ? 700 : 600,
                    display: "flex",
                    alignItems: "center",
                  },
                },
                String(value),
              ),
            ),
          ),
        ),
      ],
    ),
    { width, height },
  );

  return image.arrayBuffer();
}

async function renderDashboardHomeImage() {
  const width = 1400;
  const height = 1120;

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #08131f 0%, #0f172a 42%, #12263a 100%)",
          color: "#f8fafc",
          padding: "44px",
          fontFamily: "ui-sans-serif, system-ui, sans-serif",
          boxSizing: "border-box",
        },
      },
      [
        createElement(
          "div",
          {
            key: "hero",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              borderRadius: "34px",
              padding: "54px 42px 44px",
              background: "linear-gradient(135deg, #10251d 0%, #1f4f46 38%, #1e3a5f 100%)",
              boxShadow: "0 30px 70px rgba(2, 6, 23, 0.42)",
              border: "1px solid rgba(255,255,255,0.12)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "18px",
                  letterSpacing: "3.5px",
                  textTransform: "uppercase",
                  color: "#dbeafe",
                  textAlign: "center",
                  padding: "10px 18px",
                  borderRadius: "999px",
                  background: "rgba(255,255,255,0.08)",
                  border: "1px solid rgba(255,255,255,0.12)",
                },
              },
              "Operations Workspace",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "70px",
                  fontWeight: 800,
                  marginTop: "16px",
                  textAlign: "center",
                },
              },
              "Executive Workspace Hub",
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "26px",
                  marginTop: "14px",
                  color: "#e5eefb",
                  textAlign: "center",
                  maxWidth: "980px",
                },
              },
              "Choose the live workspace you want to open. Withdraw and Deposit stay separate, but both run in the same Telegram group.",
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "mini-stats",
            style: {
              display: "flex",
              gap: "18px",
              marginTop: "22px",
            },
          },
          [
            { label: "Mode", value: "Read-Only" },
            { label: "Workspaces", value: getAvailableDashboards().length === 2 ? "2 Active" : "1 Active" },
            { label: "Focus", value: "Withdraw + Deposit" },
          ].map((item) =>
            createElement(
              "div",
              {
                key: item.label,
                style: {
                  flex: 1,
                  display: "flex",
                  flexDirection: "column",
                  borderRadius: "24px",
                  padding: "20px 22px",
                  background: "rgba(255,255,255,0.07)",
                  border: "1px solid rgba(255,255,255,0.10)",
                  boxShadow: "0 18px 30px rgba(15, 23, 42, 0.18)",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "label",
                    style: {
                      fontSize: "18px",
                      letterSpacing: "1px",
                      textTransform: "uppercase",
                      color: "#bfdbfe",
                    },
                  },
                  item.label,
                ),
                createElement(
                  "div",
                  {
                    key: "value",
                    style: {
                      fontSize: "30px",
                      fontWeight: 800,
                      marginTop: "8px",
                    },
                  },
                  item.value,
                ),
              ],
            ),
          ),
        ),
        createElement(
          "div",
          {
            key: "cards",
            style: {
              display: "flex",
              gap: "24px",
              marginTop: "28px",
            },
          },
          [
            {
              badge: "💸",
              title: "WITHDRAW WORKSPACE",
              subtitle: "Open Real-Time, Shifting, and Daily Transactions inside the withdraw control path.",
              accent: "linear-gradient(135deg, #0f766e 0%, #1e3a5f 100%)",
            },
            {
              badge: "🏦",
              title: "DEPOSIT WORKSPACE",
              subtitle: "Open deposit sheets in their own clean dashboard so staff can switch between deposit and withdraw clearly.",
              accent: "linear-gradient(135deg, #14532d 0%, #0f766e 45%, #1e3a5f 100%)",
            },
          ].map((card) =>
            createElement(
              "div",
              {
                key: card.title,
                style: {
                  flex: 1,
                  display: "flex",
                  flexDirection: "column",
                  borderRadius: "30px",
                  padding: "28px",
                  background: card.accent,
                  boxShadow: "0 20px 40px rgba(15, 23, 42, 0.24)",
                  minHeight: "250px",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "badge",
                    style: {
                      fontSize: "54px",
                    },
                  },
                  card.badge,
                ),
                createElement(
                  "div",
                  {
                    key: "card-title",
                    style: {
                      fontSize: "34px",
                      fontWeight: 800,
                      marginTop: "18px",
                    },
                  },
                  card.title,
                ),
                createElement(
                  "div",
                  {
                    key: "card-subtitle",
                    style: {
                      fontSize: "22px",
                      marginTop: "12px",
                      lineHeight: 1.35,
                      color: "#eff6ff",
                    },
                  },
                  card.subtitle,
                ),
              ],
            ),
          ),
        ),
      ],
    ),
    { width, height },
  );

  return image.arrayBuffer();
}

async function renderRealTimeImage(preview: RealTimePreview) {
  const width = 1700;
  const height = Math.max(1540, 320 + preview.rows.length * 58);
  const columns = [
    { key: "platform", label: preview.headers[0], width: 960 },
    { key: "dayShift", label: preview.headers[1], width: 170 },
    { key: "nightShift", label: preview.headers[2], width: 170 },
    { key: "midShift", label: preview.headers[3], width: 170 },
    { key: "total", label: preview.headers[4], width: 170 },
  ] as const;

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #f7f9f7 0%, #eef3ef 100%)",
          color: "#111827",
          fontFamily: "Georgia, serif",
          boxSizing: "border-box",
          padding: "30px",
        },
      },
      [
        createElement(
          "div",
          {
            key: "stamp",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: "linear-gradient(135deg, #1f4f46 0%, #244536 55%, #163225 100%)",
              color: "#ffffff",
              padding: "22px 26px 20px",
              border: "3px solid #244536",
              borderBottom: "0",
              borderRadius: "28px 28px 0 0",
              boxShadow: "0 16px 30px rgba(15, 23, 42, 0.14)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "18px",
                  letterSpacing: "2px",
                  textTransform: "uppercase",
                  color: "#d1fae5",
                },
              },
              "Live Operations Board",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "40px",
                  fontWeight: 800,
                  marginTop: "8px",
                },
              },
              "REAL TIME",
            ),
            createElement(
              "div",
              {
                key: "timestamp",
                style: {
                  marginTop: "12px",
                  fontSize: "42px",
                  fontWeight: 700,
                  letterSpacing: "1px",
                },
              },
              preview.timestamp,
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "thead",
            style: {
              display: "flex",
              background: "#355f4d",
              color: "#ffffff",
              borderLeft: "3px solid #244536",
              borderRight: "3px solid #244536",
              borderBottom: "3px solid #1f2937",
            },
          },
          columns.map((column) =>
            createElement(
              "div",
              {
                key: column.key,
                style: {
                  width: `${column.width}px`,
                  padding: "14px 12px",
                  borderRight: column.key === "total" ? "0" : "2px solid #1f2937",
                  textAlign: "center",
                  fontSize: "28px",
                  fontWeight: 700,
                },
              },
              column.label,
            ),
          ),
        ),
        ...preview.rows.map((row, index) =>
          createElement(
            "div",
            {
              key: `${row.platform}-${index}`,
              style: {
                display: "flex",
                background: row.isTotal ? "#fde047" : index % 2 === 0 ? "#ffffff" : "#f3f7f4",
                borderLeft: "3px solid #244536",
                borderRight: "3px solid #244536",
                borderBottom: "2px solid #1f2937",
              },
            },
            columns.map((column) =>
              createElement(
                "div",
                {
                  key: `${row.platform}-${column.key}`,
                  style: {
                    width: `${column.width}px`,
                    padding: "10px 12px",
                    borderRight: column.key === "total" ? "0" : "2px solid #1f2937",
                    textAlign: column.key === "platform" ? "left" : "center",
                    fontSize: column.key === "platform" ? "30px" : "32px",
                    fontWeight: 700,
                  },
                },
                row[column.key],
              ),
            ),
          ),
        ),
      ],
    ),
    {
      width,
      height,
    },
  );

  return image.arrayBuffer();
}

async function renderMatrixSectionImage(preview: MatrixSectionPreview) {
  const columnCount = Math.max(...preview.rows.map((row) => row.length), 1);
  const suppliedWidths = preview.columnWidths ?? [];
  const fallbackWidth = Math.floor(preview.width / columnCount);
  const columnWidths = Array.from(
    { length: columnCount },
    (_, index) => suppliedWidths[index] ?? fallbackWidth,
  );
  const bodyRowCount = Math.max(preview.rows.length - preview.headerRows, 0);
  const isTeamLeaders = preview.title === "Team Leaders";
  const height = Math.max(1080, 260 + preview.headerRows * 82 + bodyRowCount * (isTeamLeaders ? 66 : 58));

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%)",
          color: "#111827",
          fontFamily: "Georgia, serif",
          boxSizing: "border-box",
          padding: "28px",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: `linear-gradient(135deg, ${preview.accent} 0%, #1f2937 100%)`,
              color: "#ffffff",
              padding: "22px 28px",
              borderRadius: "28px 28px 0 0",
              border: `3px solid ${preview.accent}`,
              borderBottom: "0",
            },
          },
          [
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "48px",
                  fontWeight: 700,
                },
              },
              `${preview.badge} ${preview.title}`,
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "24px",
                  marginTop: "8px",
                },
              },
              preview.subtitle,
            ),
          ],
        ),
        ...preview.rows.map((row, rowIndex) =>
          createElement(
            "div",
            {
              key: `row-${rowIndex}`,
              style: {
                display: "flex",
                background:
                  rowIndex < preview.headerRows
                    ? preview.accent
                    : rowIndex % 2 === 0
                      ? "#ffffff"
                      : "#f7f8fc",
                color: rowIndex < preview.headerRows ? "#ffffff" : "#111827",
                borderLeft: `3px solid ${preview.accent}`,
                borderRight: `3px solid ${preview.accent}`,
                borderBottom: "2px solid #cbd5e1",
              },
            },
            Array.from({ length: columnCount }, (_, columnIndex) =>
              createElement(
                "div",
                {
                  key: `cell-${rowIndex}-${columnIndex}`,
                  style: {
                    width: `${columnWidths[columnIndex]}px`,
                    padding: rowIndex < preview.headerRows ? "12px 10px" : "10px 10px",
                    borderRight: columnIndex === columnCount - 1 ? "0" : "2px solid #1f2937",
                    textAlign: columnIndex === 0 ? "left" : "center",
                    fontSize: rowIndex < preview.headerRows ? "26px" : isTeamLeaders ? "21px" : "24px",
                    fontWeight: rowIndex < preview.headerRows ? 700 : columnIndex === 0 ? 700 : 600,
                    whiteSpace: "pre-wrap",
                    wordBreak: "break-word",
                    lineHeight: 1.15,
                    minHeight: rowIndex < preview.headerRows ? "64px" : isTeamLeaders ? "60px" : "52px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: columnIndex === 0 ? "flex-start" : "center",
                  },
                },
                row[columnIndex] ?? "",
              ),
            ),
          ),
        ),
      ],
    ),
    {
      width: preview.width,
      height,
    },
  );

  return image.arrayBuffer();
}

function renderShiftingEntryCard(entry: ShiftingEntry, accent: string) {
  return createElement(
    "div",
    {
      key: `${entry.name}-${entry.id}-${entry.shift}`,
      style: {
        display: "flex",
        flexDirection: "column",
        gap: "10px",
        padding: "20px 24px",
        borderRadius: "24px",
        background: "rgba(255,255,255,0.98)",
        border: `2px solid ${accent}`,
        boxShadow: "0 12px 28px rgba(15, 23, 42, 0.10)",
        marginBottom: "16px",
      },
    },
    [
      createElement(
        "div",
        {
          key: "top",
          style: {
            display: "flex",
            justifyContent: "center",
            flexWrap: "wrap",
            gap: "14px",
            fontSize: "18px",
            alignItems: "center",
          },
        },
        [
          createElement(
            "div",
            {
              key: "shift",
              style: {
                color: accent,
                fontWeight: 700,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                textAlign: "center",
                padding: "8px 14px",
                borderRadius: "999px",
                background: `${accent}18`,
              },
            },
            entry.shift || "Shift",
          ),
          createElement(
            "div",
            {
              key: "role",
              style: {
                color: "#475569",
                fontSize: "17px",
                textAlign: "center",
                padding: "8px 14px",
                borderRadius: "999px",
                background: "rgba(148, 163, 184, 0.12)",
              },
            },
            entry.role || "Role",
          ),
        ],
      ),
      createElement(
        "div",
        {
          key: "name",
          style: {
            fontSize: "30px",
            fontWeight: 800,
            color: "#0f172a",
            textAlign: "center",
            lineHeight: 1.2,
          },
        },
        shortenCell(entry.name, 36),
      ),
      createElement(
        "div",
        {
          key: "meta",
          style: {
            fontSize: "18px",
            color: "#334155",
            textAlign: "center",
            fontWeight: 600,
          },
        },
        `${entry.platform || "Platform"} | ${entry.id || "No ID"}`,
      ),
      createElement(
        "div",
        {
          key: "account",
          style: {
            fontSize: "16px",
            color: "#64748b",
            textAlign: "center",
            lineHeight: 1.3,
          },
        },
        shortenCell(entry.account || entry.startDate || "", 44),
      ),
    ],
  );
}

async function renderShiftingOverviewImage(
  summary: ShiftingSummaryItem[],
  platformCount: number,
) {
  const width = 1400;
  const summaryAccent = ["#1d4ed8", "#0f766e", "#9333ea", "#c2410c"];
  const summaryRows = Math.ceil(Math.max(summary.length, 1) / 4);
  const height = Math.max(900, 360 + summaryRows * 150);

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #eef6f4 0%, #e2ecf6 100%)",
          color: "#0f172a",
          padding: "34px",
          fontFamily: "ui-sans-serif, system-ui, sans-serif",
          boxSizing: "border-box",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: "linear-gradient(135deg, #1f4f46 0%, #24584e 100%)",
              color: "#ffffff",
              borderRadius: "30px",
              padding: "32px 34px",
              marginBottom: "26px",
              boxShadow: "0 18px 36px rgba(15, 23, 42, 0.18)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "22px",
                  letterSpacing: "2px",
                  textTransform: "uppercase",
                  color: "#c7f9cc",
                  textAlign: "center",
                },
              },
              "Team Shifting Board",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "52px",
                  fontWeight: 800,
                  marginTop: "10px",
                  textAlign: "center",
                },
              },
              "SHIFTING",
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "26px",
                  color: "#dbeafe",
                  marginTop: "10px",
                  textAlign: "center",
                },
              },
              `${platformCount} platform${platformCount === 1 ? "" : "s"} available · control center ready`,
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "summary",
            style: {
              display: "flex",
              flexWrap: "wrap",
              gap: "16px",
              justifyContent: "center",
            },
          },
          summary.map((item, index) =>
            createElement(
              "div",
              {
                key: item.label,
                style: {
                  width: "315px",
                  display: "flex",
                  flexDirection: "column",
                  padding: "20px 22px",
                  borderRadius: "20px",
                  background: "#ffffff",
                  borderTop: `6px solid ${summaryAccent[index % summaryAccent.length]}`,
                  boxShadow: "0 10px 26px rgba(15, 23, 42, 0.10)",
                  alignItems: "center",
                  justifyContent: "center",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "label",
                    style: {
                      fontSize: "20px",
                      color: "#475569",
                      textAlign: "center",
                    },
                  },
                  item.label,
                ),
                createElement(
                  "div",
                  {
                    key: "value",
                    style: {
                      fontSize: "44px",
                      fontWeight: 800,
                      color: "#0f172a",
                      marginTop: "10px",
                      textAlign: "center",
                    },
                  },
                  item.value,
                ),
              ],
            ),
          ),
        ),
      ],
    ),
    { width, height },
  );

  return image.arrayBuffer();
}

async function renderShiftingImage(preview: ShiftingPreview) {
  const width = 1400;
  const shiftAccent = preview.shiftKind === "night" ? "#1d4ed8" : preview.shiftKind === "mid" ? "#7c3aed" : "#0f766e";
  const shiftLabel = preview.shiftKind === "night" ? "Night Shift" : preview.shiftKind === "mid" ? "Mid Shift" : "Day Shift";
  const shiftBadge = preview.shiftKind === "night" ? "🌙" : preview.shiftKind === "mid" ? "🌇" : "🌤️";
  const pageLabel = preview.totalEntryPages > 1 ? `Showing ${preview.pagedEntries.length} of ${preview.entries.length} · Page ${preview.entryPage + 1} of ${preview.totalEntryPages}` : `Showing ${preview.pagedEntries.length} of ${preview.entries.length}`;
  const entryCount = Math.max(preview.pagedEntries.length, 1);
  const height = Math.max(980, 280 + entryCount * 156);

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #eef6f4 0%, #e2ecf6 100%)",
          color: "#0f172a",
          padding: "34px",
          fontFamily: "ui-sans-serif, system-ui, sans-serif",
          boxSizing: "border-box",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: `linear-gradient(135deg, ${shiftAccent} 0%, #1f4f46 100%)`,
              color: "#ffffff",
              borderRadius: "30px",
              padding: "28px 30px",
              marginBottom: "22px",
              boxShadow: "0 18px 36px rgba(15, 23, 42, 0.18)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "22px",
                  letterSpacing: "2px",
                  textTransform: "uppercase",
                  color: "#c7f9cc",
                  textAlign: "center",
                },
              },
              "Shifting Command Center",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "48px",
                  fontWeight: 800,
                  marginTop: "8px",
                  textAlign: "center",
                },
              },
              preview.currentSection.title,
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "24px",
                  color: "#dbeafe",
                  marginTop: "8px",
                  textAlign: "center",
                },
              },
              `${shiftBadge} ${shiftLabel} • ${pageLabel}`,
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "list-card",
            style: {
              display: "flex",
              flexDirection: "column",
              flex: 1,
              background: "rgba(255,255,255,0.82)",
              borderRadius: "30px",
              padding: "26px",
              boxShadow: "0 18px 36px rgba(15, 23, 42, 0.10)",
            },
          },
          [
            ...(preview.pagedEntries.length > 0
              ? preview.pagedEntries.map((entry) => renderShiftingEntryCard(entry, shiftAccent))
              : [
                  createElement(
                    "div",
                    {
                      key: "empty",
                      style: {
                        fontSize: "22px",
                        color: "#64748b",
                        padding: "18px 8px",
                      },
                    },
                    `No ${preview.shiftKind === "night" ? "night" : preview.shiftKind === "mid" ? "mid" : "day"} entries in this platform.`,
                  ),
                ]),
          ],
        ),
      ],
    ),
    {
      width,
      height,
    },
  );

  return image.arrayBuffer();
}

async function renderDailyTransactionOverviewImage(summary: DailyTransactionSummary) {
  const width = 1400;
  const height = 920;
  const cards = [
    { label: "Platforms", value: String(summary.totalPlatforms), accent: "#1d4ed8" },
    { label: "Staff Profiles", value: String(summary.totalStaff), accent: "#0f766e" },
    { label: "Day Shift", value: String(summary.dayCount), accent: "#ea580c" },
    { label: "Mid Shift", value: String(summary.midCount), accent: "#7c3aed" },
    { label: "Night Shift", value: String(summary.nightCount), accent: "#1e40af" },
  ];

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #edf6ff 0%, #e8f3ee 100%)",
          color: "#0f172a",
          padding: "34px",
          fontFamily: "ui-sans-serif, system-ui, sans-serif",
          boxSizing: "border-box",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: "linear-gradient(135deg, #14532d 0%, #1f4f46 45%, #1e3a5f 100%)",
              color: "#ffffff",
              borderRadius: "30px",
              padding: "34px 38px",
              marginBottom: "28px",
              boxShadow: "0 18px 36px rgba(15, 23, 42, 0.18)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "22px",
                  letterSpacing: "2px",
                  textTransform: "uppercase",
                  color: "#d1fae5",
                },
              },
              "Daily Transactions Board",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "54px",
                  fontWeight: 800,
                  marginTop: "10px",
                  textAlign: "center",
                },
              },
              "APRIL DAILY TRANSACTIONS",
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "26px",
                  color: "#dbeafe",
                  marginTop: "10px",
                  textAlign: "center",
                },
              },
              "Choose a platform, then a shift, to review all staff transaction activity together.",
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "cards",
            style: {
              display: "flex",
              gap: "16px",
              flexWrap: "wrap",
              justifyContent: "center",
            },
          },
          cards.map((card) =>
            createElement(
              "div",
              {
                key: card.label,
                style: {
                  width: "240px",
                  display: "flex",
                  flexDirection: "column",
                  padding: "20px 22px",
                  borderRadius: "22px",
                  background: "#ffffff",
                  borderTop: `6px solid ${card.accent}`,
                  boxShadow: "0 12px 28px rgba(15, 23, 42, 0.10)",
                  alignItems: "center",
                  justifyContent: "center",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "label",
                    style: {
                      fontSize: "22px",
                      color: "#475569",
                      textAlign: "center",
                    },
                  },
                  card.label,
                ),
                createElement(
                  "div",
                  {
                    key: "value",
                    style: {
                      fontSize: "46px",
                      fontWeight: 800,
                      marginTop: "10px",
                    },
                  },
                  card.value,
                ),
              ],
            ),
          ),
        ),
      ],
    ),
    { width, height },
  );

  return image.arrayBuffer();
}

async function renderDailyTransactionEntryImage(preview: DailyTransactionPreview) {
  const width = 1400;
  const normalizedShift = preview.shiftKind;
  const height = Math.max(1040, 360 + Math.max(preview.pagedEntries.length, 1) * 220);
  const shiftAccent =
    normalizedShift === "night"
      ? "#1d4ed8"
      : normalizedShift === "mid"
        ? "#7c3aed"
        : "#0f766e";
  const shiftLabel =
    normalizedShift === "night"
      ? "🌙 Night Shift / 夜班"
      : normalizedShift === "mid"
        ? "🌇 Mid Shift / 中班"
        : "🌤️ Day Shift / 白班";
  const overallStats = getTransactionStats(preview.entries.flatMap((entry) => entry.metrics));

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #eef6f4 0%, #e2ecf6 100%)",
          color: "#0f172a",
          padding: "34px",
          fontFamily: "ui-sans-serif, system-ui, sans-serif",
          boxSizing: "border-box",
        },
      },
      [
        createElement(
          "div",
          {
            key: "header",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              background: `linear-gradient(135deg, ${shiftAccent} 0%, #1f4f46 100%)`,
              color: "#ffffff",
              borderRadius: "30px",
              padding: "30px 34px",
              marginBottom: "24px",
              boxShadow: "0 18px 36px rgba(15, 23, 42, 0.18)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "20px",
                  letterSpacing: "2px",
                  textTransform: "uppercase",
                  color: "#dbeafe",
                },
              },
              "Daily Transactions Board",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "48px",
                  fontWeight: 800,
                  marginTop: "12px",
                  textAlign: "center",
                },
              },
              preview.currentPlatform.title,
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "24px",
                  color: "#e2e8f0",
                  marginTop: "10px",
                  textAlign: "center",
                },
              },
              `${shiftLabel} · ${preview.entries.length} staff profile${preview.entries.length === 1 ? "" : "s"}`,
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "summary-cards",
            style: {
              display: "flex",
              gap: "18px",
              justifyContent: "center",
              marginBottom: "20px",
            },
          },
          [
            { label: "Staff Count", value: String(preview.entries.length), accent: "#0f766e" },
            { label: "RD", value: String(overallStats.rdCount), accent: "#ea580c" },
            { label: "ABS", value: String(overallStats.absCount), accent: "#dc2626" },
            { label: "Peak Value", value: overallStats.peakValue > 0 ? overallStats.peakValue.toLocaleString("en-US") : "-", accent: shiftAccent },
            { label: "Total Processed", value: overallStats.totalProcessed > 0 ? overallStats.totalProcessed.toLocaleString("en-US") : "-", accent: "#1e3a5f" },
          ].map((item) =>
            createElement(
              "div",
              {
                key: item.label,
                style: {
                  flex: 1,
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "center",
                  justifyContent: "center",
                  padding: "18px 16px",
                  borderRadius: "22px",
                  background: "#ffffff",
                  borderTop: `6px solid ${item.accent}`,
                  boxShadow: "0 10px 24px rgba(15, 23, 42, 0.10)",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "label",
                    style: {
                      fontSize: "16px",
                      textTransform: "uppercase",
                      letterSpacing: "1px",
                      color: "#64748b",
                      textAlign: "center",
                    },
                  },
                  item.label,
                ),
                createElement(
                  "div",
                  {
                    key: "value",
                    style: {
                      fontSize: "28px",
                      fontWeight: 800,
                      color: "#0f172a",
                      marginTop: "8px",
                      textAlign: "center",
                    },
                  },
                  item.value,
                ),
              ],
            ),
          ),
        ),
        createElement(
          "div",
          {
            key: "meta",
            style: {
              display: "flex",
              gap: "18px",
              justifyContent: "center",
              marginBottom: "20px",
            },
          },
          [
            { label: "Platform", value: preview.currentPlatform.title || "-" },
            { label: "View", value: `${shiftLabel} · Page ${preview.entryPage + 1} / ${preview.totalEntryPages}` },
          ].map((item) =>
            createElement(
              "div",
              {
                key: item.label,
                style: {
                  flex: 1,
                  display: "flex",
                  flexDirection: "column",
                  padding: "18px 22px",
                  borderRadius: "22px",
                  background: "#ffffff",
                  boxShadow: "0 10px 24px rgba(15, 23, 42, 0.10)",
                  maxWidth: "620px",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "label",
                    style: {
                      fontSize: "18px",
                      textTransform: "uppercase",
                      letterSpacing: "1px",
                      color: "#64748b",
                    },
                  },
                  item.label,
                ),
                createElement(
                  "div",
                  {
                    key: "value",
                    style: {
                      fontSize: "30px",
                      fontWeight: 800,
                      marginTop: "8px",
                      color: "#0f172a",
                      wordBreak: "break-word",
                    },
                  },
                  item.value,
                ),
              ],
            ),
          ),
        ),
        createElement(
          "div",
          {
            key: "metrics",
            style: {
              display: "flex",
              flexDirection: "column",
              gap: "12px",
            },
          },
          preview.pagedEntries.length > 0
            ? preview.pagedEntries.map((entry) => {
                const entryStats = getTransactionStats(entry.metrics);

                return createElement(
                  "div",
                  {
                    key: `${entry.name}-${entry.systemUser}`,
                    style: {
                      display: "flex",
                      flexDirection: "column",
                      gap: "12px",
                      padding: "20px 24px",
                      borderRadius: "22px",
                      background: "#ffffff",
                      border: `2px solid ${shiftAccent}`,
                      boxShadow: "0 10px 24px rgba(15, 23, 42, 0.08)",
                    },
                  },
                  [
                    createElement(
                      "div",
                      {
                        key: "top",
                        style: {
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          gap: "12px",
                        },
                      },
                      [
                        createElement(
                          "div",
                          {
                            key: "name",
                            style: {
                              fontSize: "28px",
                              fontWeight: 800,
                              color: "#0f172a",
                            },
                          },
                          entry.name,
                        ),
                        createElement(
                          "div",
                          {
                            key: "user",
                            style: {
                              fontSize: "18px",
                              color: "#475569",
                              fontWeight: 700,
                            },
                          },
                          entry.systemUser || "-",
                        ),
                      ],
                    ),
                    createElement(
                      "div",
                      {
                        key: "entry-summary",
                        style: {
                          display: "flex",
                          gap: "10px",
                          flexWrap: "wrap",
                        },
                      },
                      [
                        { label: "Processed", value: entryStats.totalProcessed > 0 ? entryStats.totalProcessed.toLocaleString("en-US") : "-" },
                        { label: "Peak", value: entryStats.peakValue > 0 ? entryStats.peakValue.toLocaleString("en-US") : "-" },
                        { label: "Active Days", value: String(entryStats.activeDays) },
                        { label: "RD", value: String(entryStats.rdCount) },
                        { label: "ABS", value: String(entryStats.absCount) },
                      ].map((item) =>
                        createElement(
                          "div",
                          {
                            key: item.label,
                            style: {
                              display: "flex",
                              flexDirection: "column",
                              alignItems: "center",
                              justifyContent: "center",
                              minWidth: "110px",
                              padding: "10px 12px",
                              borderRadius: "16px",
                              background: "#f8fafc",
                              border: "1px solid #cbd5e1",
                            },
                          },
                          [
                            createElement(
                              "div",
                              {
                                key: "label",
                                style: {
                                  fontSize: "13px",
                                  color: "#64748b",
                                  textTransform: "uppercase",
                                  letterSpacing: "0.5px",
                                },
                              },
                              item.label,
                            ),
                            createElement(
                              "div",
                              {
                                key: "value",
                                style: {
                                  marginTop: "4px",
                                  fontSize: "20px",
                                  fontWeight: 800,
                                  color: "#0f172a",
                                },
                              },
                              item.value,
                            ),
                          ],
                        ),
                      ),
                    ),
                    createElement(
                      "div",
                      {
                        key: "grid",
                        style: {
                          display: "flex",
                          flexWrap: "wrap",
                          gap: "10px",
                        },
                      },
                      entry.metrics.map((metric) =>
                        createElement(
                          "div",
                          {
                            key: `${entry.name}-${metric.label}`,
                            style: {
                              width: "120px",
                              display: "flex",
                              flexDirection: "column",
                              alignItems: "center",
                              justifyContent: "center",
                              padding: "12px 8px",
                              borderRadius: "18px",
                              background:
                                metric.value.toUpperCase() === "RD"
                                  ? "linear-gradient(135deg, #fff7ed 0%, #ffedd5 100%)"
                                  : metric.value.toUpperCase() === "ABS"
                                    ? "linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%)"
                                    : "#f8fafc",
                              border:
                                metric.value.toUpperCase() === "RD"
                                  ? "2px solid #f97316"
                                  : metric.value.toUpperCase() === "ABS"
                                    ? "2px solid #ef4444"
                                    : "1px solid #cbd5e1",
                            },
                          },
                          [
                            createElement(
                              "div",
                              {
                                key: "label",
                                style: {
                                  fontSize: "16px",
                                  color: "#64748b",
                                  fontWeight: 700,
                                },
                              },
                              `Day ${metric.label}`,
                            ),
                            createElement(
                              "div",
                              {
                                key: "value",
                                style: {
                                  marginTop: "6px",
                                  fontSize: "24px",
                                  fontWeight: 800,
                                  color:
                                    metric.value.toUpperCase() === "RD"
                                      ? "#c2410c"
                                      : metric.value.toUpperCase() === "ABS"
                                        ? "#b91c1c"
                                        : "#0f172a",
                                },
                              },
                              metric.value || "-",
                            ),
                          ],
                        ),
                      ),
                    ),
                  ],
                );
              })
            : [
                createElement(
                  "div",
                  {
                    key: "empty",
                    style: {
                      fontSize: "24px",
                      color: "#64748b",
                      padding: "26px 18px",
                      textAlign: "center",
                      background: "#ffffff",
                      borderRadius: "22px",
                      border: `2px dashed ${shiftAccent}`,
                    },
                  },
                  `No ${shiftLabel.toLowerCase()} entries found in ${preview.currentPlatform.title}.`,
                ),
              ],
        ),
      ],
    ),
    { width, height },
  );

  return image.arrayBuffer();
}

async function sendHomeMenu(chatId: number, text?: string) {
  const replyMarkup = await buildHomeKeyboard();
  const imageBuffer = await renderDashboardHomeImage();

  await sendTelegramPhoto(
    chatId,
    imageBuffer,
    text ?? `OPERATIONS HUB
Executive Workspace Hub
Choose Withdraw or Deposit.`,
    replyMarkup,
  );
}

async function sendDashboardMenu(chatId: number, dashboard: DashboardKey, text?: string) {
  const replyMarkup = await buildSheetKeyboard(dashboard);

  await sendKeyboardMessage(
    chatId,
    text ??
      `${getDashboardBadge(dashboard)} ${getDashboardLabel(dashboard)}
Choose a board below to continue.`,
    replyMarkup,
  );
}

async function showSheet(
  callbackQuery: TelegramCallbackQuery,
  dashboard: DashboardKey,
  sheetIndex: number,
  page: number,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "The original message is no longer available.");
    return;
  }

  const sheets = await getDashboardSheets(dashboard);
  const sheet = sheets[sheetIndex];

  if (!sheet) {
    await answerCallbackQuery(callbackQuery.id, "That sheet was not found.");
    return;
  }

  await answerCallbackQuery(callbackQuery.id, "Loading...");

  const loadingMessageId = await sendStatusMessage(message.chat.id, getLoadingMessage(sheet.title));

  try {
    let imageBuffer: ArrayBuffer;
    let caption: string;
    let replyMarkup: SheetNavigation;

    if (dashboard === "withdraw" && sheet.title === "REAL TIME") {
      const safePage = Math.max(0, Math.min(page, REAL_TIME_SECTION_COUNT - 1));

      if (safePage === 0) {
        const preview = await getRealTimeSummaryPreview();
        imageBuffer = await renderRealTimeImage(preview);
        caption = buildRealTimeSummaryCaption(preview);
      } else {
        const preview = await getRealTimeMatrixSection(safePage);
        imageBuffer = await renderMatrixSectionImage(preview);
        caption = buildRealTimeSectionCaption(preview);
      }

      replyMarkup = buildSectionNavigation(dashboard, sheetIndex, safePage, REAL_TIME_SECTION_COUNT);
    } else if (dashboard === "withdraw" && sheet.title === "BASICS WITHDARW") {
      const preview = await getBasicsWithdrawPreview();
      imageBuffer = await renderBasicsWithdrawImage(preview);
      caption = buildBasicsCaption();
      replyMarkup = {
        inline_keyboard: [
          [{ text: "↩ Withdraw Dashboard", callback_data: "menu:withdraw" }],
          [{ text: "🧭 Workspaces", callback_data: "home:0" }],
        ],
      };
    } else if (dashboard === "withdraw" && sheet.title === "WORKFOLIO EMAIL") {
      const preview = await getWorkfolioEmailPreview(page);
      imageBuffer = await renderWorkfolioEmailImage(preview);
      caption = buildWorkfolioEmailCaption(preview);
      replyMarkup = buildSectionNavigation(dashboard, sheetIndex, preview.page, preview.sections.length);
    } else if (dashboard === "withdraw" && sheet.title === "SHIFTING") {
      const preview = await getShiftingPreview(page, "day");
      imageBuffer = await renderShiftingImage(preview);
      caption = buildShiftingCaption(preview);
      replyMarkup = buildShiftingShiftKeyboard(
        sheetIndex,
        preview.platformIndex,
        preview.currentSection,
        preview.shiftKind,
        preview.sections.length,
        preview.entryPage,
        preview.totalEntryPages,
      );
    } else {
      const config = getSheetWindowConfig(sheet.title);
      const window = await withTimeout(
        readSheetWindow(
          sheet.title,
          page,
          config.rowsPerPage,
          config.columnsToShow,
          getSpreadsheetIdForDashboard(dashboard),
        ),
        12000,
        `${sheet.title} data load`,
      );
      imageBuffer = await withTimeout(
        renderGenericSheetImage(
          sheet.title,
          window.rowOffset,
          sheet.rowCount,
          window.rows,
        ),
        12000,
        `${sheet.title} image render`,
      );
      caption = buildSheetCaption(sheet.title, window.rowOffset, sheet.rowCount, window.rows);
      const lastVisibleRow = window.rows.length > 0 ? window.rowOffset + window.rows.length - 1 : window.rowOffset;
      const totalPages = Math.max(1, Math.ceil(sheet.rowCount / config.rowsPerPage));
      replyMarkup = buildSheetNavigation(
        dashboard,
        sheetIndex,
        page,
        totalPages,
        `Rows ${window.rowOffset}-${lastVisibleRow}`,
      );
    }

    if (message.photo && message.photo.length > 0) {
      await deleteTelegramMessage(message.chat.id, message.message_id);
    }

    await sendTelegramPhoto(message.chat.id, imageBuffer, caption, replyMarkup);
  } catch (error) {
    console.error(`Failed to render sheet ${sheet.title}`, error);
    await callTelegram("sendMessage", {
      chat_id: message.chat.id,
      text: `⚠️ ${sheet.title} is taking too long right now. Please tap it again in a moment.`,
      reply_markup: {
        inline_keyboard: [[{ text: `↩ ${getDashboardLabel(dashboard)}`, callback_data: `menu:${dashboard}` }]],
      },
    }).catch(() => undefined);
  } finally {
    if (loadingMessageId) {
      await deleteTelegramMessage(message.chat.id, loadingMessageId).catch(() => undefined);
    }
  }
}

async function handleMessage(message: TelegramMessage) {
  if (!isAllowedChat(message.chat.id)) {
    await callTelegram("sendMessage", {
      chat_id: message.chat.id,
      text: "This bot is private. Ask the admin to allow this Telegram chat first.",
    });
    return;
  }

  const text = message.text?.trim().toLowerCase() ?? "";

  if (text === "/start" || text === "/menu") {
    await sendHomeMenu(message.chat.id);
    return;
  }

  if (text === "/help") {
    await callTelegram("sendMessage", {
      chat_id: message.chat.id,
      text: `Commands:
/start - open sheet menu
/menu - open workspace menu
/help - show commands
/chatid - show this chat ID

This bot is read-only and supports separate Withdraw and Deposit workspaces in the same group.`,
    });
    return;
  }

  if (text === "/chatid") {
    await callTelegram("sendMessage", {
      chat_id: message.chat.id,
      text: `This chat ID is: ${message.chat.id}`,
    });
    return;
  }

  await callTelegram("sendMessage", {
    chat_id: message.chat.id,
    text: "Use /start to open the sheet menu.",
  });
}

async function handleCallbackQuery(callbackQuery: TelegramCallbackQuery) {
  const message = callbackQuery.message;
  const chatId = message?.chat.id;

  if (!chatId || !message) {
    await answerCallbackQuery(callbackQuery.id, "Missing chat context.");
    return;
  }

  if (!isAllowedChat(chatId)) {
    await answerCallbackQuery(callbackQuery.id, "This bot is private.");
    return;
  }

  const data = callbackQuery.data ?? "";

  if (data.startsWith("noop:")) {
    await answerCallbackQuery(callbackQuery.id);
    return;
  }

  if (data.startsWith("menu:")) {
    await answerCallbackQuery(callbackQuery.id);
    const [, dashboardValue] = data.split(":");

    if (message.photo && message.photo.length > 0) {
      await deleteTelegramMessage(chatId, message.message_id);
    }

    if (dashboardValue === "withdraw" || dashboardValue === "deposit") {
      await sendDashboardMenu(
        chatId,
        dashboardValue,
        `${getDashboardBadge(dashboardValue)} ${getDashboardLabel(dashboardValue)}
Choose a board below to continue.`,
      );
      return;
    }

    await sendHomeMenu(chatId);
    return;
  }

  if (data.startsWith("home:")) {
    await answerCallbackQuery(callbackQuery.id);

    if (message.photo && message.photo.length > 0) {
      await deleteTelegramMessage(chatId, message.message_id);
    }

    await sendHomeMenu(chatId);
    return;
  }

  if (data.startsWith("txplatforms:")) {
    const [, sheetIndexValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);

    if (Number.isNaN(sheetIndex)) {
      await answerCallbackQuery(callbackQuery.id, "Invalid transactions request.");
      return;
    }

    await showDailyTransactionPlatformMenu(callbackQuery, sheetIndex);
    return;
  }

  if (data.startsWith("txplatform:")) {
    const [, sheetIndexValue, platformIndexValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);
    const platformIndex = Number(platformIndexValue);

    if (Number.isNaN(sheetIndex) || Number.isNaN(platformIndex)) {
      await answerCallbackQuery(callbackQuery.id, "Invalid platform request.");
      return;
    }

    await showDailyTransactionShiftMenu(callbackQuery, sheetIndex, platformIndex);
    return;
  }

  if (data.startsWith("txview:")) {
    const [, sheetIndexValue, platformIndexValue, shiftKindValue, entryPageValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);
    const platformIndex = Number(platformIndexValue);
    const entryPage = Number(entryPageValue ?? "0");

    if (
      Number.isNaN(sheetIndex) ||
      Number.isNaN(platformIndex) ||
      Number.isNaN(entryPage) ||
      (shiftKindValue !== "day" && shiftKindValue !== "mid" && shiftKindValue !== "night")
    ) {
      await answerCallbackQuery(callbackQuery.id, "Invalid transactions request.");
      return;
    }

    await showDailyTransactionView(callbackQuery, sheetIndex, platformIndex, shiftKindValue, entryPage);
    return;
  }

  if (data.startsWith("shiftplatforms:")) {
    const [, sheetIndexValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);

    if (Number.isNaN(sheetIndex)) {
      await answerCallbackQuery(callbackQuery.id, "Invalid SHIFTING request.");
      return;
    }

    await showShiftingPlatformMenu(callbackQuery, sheetIndex);
    return;
  }

  if (data.startsWith("shiftplatform:")) {
    const [, sheetIndexValue, platformIndexValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);
    const platformIndex = Number(platformIndexValue);

    if (Number.isNaN(sheetIndex) || Number.isNaN(platformIndex)) {
      await answerCallbackQuery(callbackQuery.id, "Invalid platform request.");
      return;
    }

    await showShiftingShiftMenu(callbackQuery, sheetIndex, platformIndex);
    return;
  }

  if (data.startsWith("shiftview:")) {
    const [, sheetIndexValue, platformIndexValue, shiftKindValue, entryPageValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);
    const platformIndex = Number(platformIndexValue);
    const entryPage = Number(entryPageValue ?? "0");

    if (
      Number.isNaN(sheetIndex) ||
      Number.isNaN(platformIndex) ||
      Number.isNaN(entryPage) ||
      (shiftKindValue !== "day" && shiftKindValue !== "mid" && shiftKindValue !== "night")
    ) {
      await answerCallbackQuery(callbackQuery.id, "Invalid shift request.");
      return;
    }

    await showShiftingView(callbackQuery, sheetIndex, platformIndex, shiftKindValue, entryPage);
    return;
  }

  if (data.startsWith("sheet:")) {
    const [, dashboardValue, sheetIndexValue, pageValue] = data.split(":");
    const dashboard = dashboardValue === "deposit" ? "deposit" : "withdraw";
    const sheetIndex = Number(sheetIndexValue);
    const page = Number(pageValue);

    if (Number.isNaN(sheetIndex) || Number.isNaN(page)) {
      await answerCallbackQuery(callbackQuery.id, "Invalid sheet request.");
      return;
    }

    const sheets = await getDashboardSheets(dashboard);
    const sheet = sheets[sheetIndex];

    if (dashboard === "withdraw" && sheet?.title === "SHIFTING") {
      await showShiftingPlatformMenu(callbackQuery, sheetIndex);
      return;
    }

    if (dashboard === "withdraw" && sheet?.title === "APRIL DAILY TRANSACTIONS") {
      await showDailyTransactionPlatformMenu(callbackQuery, sheetIndex);
      return;
    }

    await showSheet(callbackQuery, dashboard, sheetIndex, page);
    return;
  }

  await answerCallbackQuery(callbackQuery.id, "Unknown action.");
}

export async function GET() {
  return NextResponse.json({
    ok: true,
    message: "Telegram webhook route is ready.",
  });
}

export async function POST(request: Request) {
  try {
    const update = (await request.json()) as TelegramUpdate;

    if (update.message) {
      await handleMessage(update.message);
    }

    if (update.callback_query) {
      await handleCallbackQuery(update.callback_query);
    }

    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error("Telegram webhook failed", error);

    return NextResponse.json(
      {
        ok: false,
        message:
          error instanceof Error
            ? error.message
            : "Telegram webhook processing failed.",
      },
      { status: 500 },
    );
  }
}

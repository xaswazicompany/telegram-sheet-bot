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

const DASHBOARD_SHEET_TITLES = ["REAL TIME", "SHIFTING"] as const;

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
  const entriesPerPage = 4;
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
  const pageLabel = preview.totalEntryPages > 1 ? `
Member Page ${preview.entryPage + 1} of ${preview.totalEntryPages}` : "";

  return `🔄 SHIFTING
${preview.currentSection.title}
${shiftLabel} · ${preview.entries.length} team member${preview.entries.length === 1 ? "" : "s"}${pageLabel}`;
}

function buildSheetNavigation(sheetIndex: number, page: number, totalPages: number, rowLabel?: string): SheetNavigation {
  const inlineKeyboard: InlineKeyboardButton[][] = [];
  const navigationButtons: InlineKeyboardButton[] = [];
  const safeTotalPages = Math.max(1, totalPages);
  const lastPage = safeTotalPages - 1;

  if (page > 0) {
    navigationButtons.push({
      text: "⏮️ First",
      callback_data: `sheet:${sheetIndex}:0`,
    });
    navigationButtons.push({
      text: "⬅️ Previous",
      callback_data: `sheet:${sheetIndex}:${page - 1}`,
    });
  }

  if (page < lastPage) {
    navigationButtons.push({
      text: "Next ➡️",
      callback_data: `sheet:${sheetIndex}:${page + 1}`,
    });
    navigationButtons.push({
      text: "Last ⏭️",
      callback_data: `sheet:${sheetIndex}:${lastPage}`,
    });
  }

  if (navigationButtons.length > 0) {
    inlineKeyboard.push(navigationButtons);
  }

  const jumpButtons: InlineKeyboardButton[] = [];

  if (page >= 10) {
    jumpButtons.push({
      text: "⏪ -10",
      callback_data: `sheet:${sheetIndex}:${Math.max(0, page - 10)}`,
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
      callback_data: `sheet:${sheetIndex}:${Math.min(lastPage, page + 10)}`,
    });
  }

  if (jumpButtons.length > 0) {
    inlineKeyboard.push(jumpButtons);
  }

  inlineKeyboard.push([{ text: "🏠 Dashboard", callback_data: "menu:0" }]);

  return { inline_keyboard: inlineKeyboard };
}

function buildSectionNavigation(sheetIndex: number, page: number, totalSections: number): SheetNavigation {
  const inlineKeyboard: InlineKeyboardButton[][] = [];
  const navigationButtons: InlineKeyboardButton[] = [];

  if (page > 0) {
    navigationButtons.push({
      text: "⬅️ Previous",
      callback_data: `sheet:${sheetIndex}:${page - 1}`,
    });
  }

  if (page + 1 < totalSections) {
    navigationButtons.push({
      text: "Next ➡️",
      callback_data: `sheet:${sheetIndex}:${page + 1}`,
    });
  }

  if (navigationButtons.length > 0) {
    inlineKeyboard.push(navigationButtons);
  }

  inlineKeyboard.push([{ text: "🏠 Dashboard", callback_data: "menu:0" }]);

  return { inline_keyboard: inlineKeyboard };
}

function buildShiftingPlatformKeyboard(sheetIndex: number, sections: ShiftingSection[]): SheetNavigation {
  const rows: InlineKeyboardButton[][] = [];

  for (let index = 0; index < sections.length; index += 2) {
    rows.push(
      sections.slice(index, index + 2).map((section, offset) => ({
        text: shortenCell(section.title, 24),
        callback_data: `shiftplatform:${sheetIndex}:${index + offset}`,
      })),
    );
  }

  rows.push([{ text: "🏠 Dashboard", callback_data: "menu:0" }]);

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
      text: "⬅️ Prev Platform",
      callback_data: `shiftview:${sheetIndex}:${platformIndex - 1}:${shiftKind}:0`,
    });
  }

  if (platformIndex + 1 < totalPlatforms) {
    platformNav.push({
      text: "Next Platform ➡️",
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
        text: "⬅️ Prev Members",
        callback_data: `shiftview:${sheetIndex}:${platformIndex}:${shiftKind}:${entryPage - 1}`,
      });
    }

    if (entryPage + 1 < totalEntryPages) {
      pageNav.push({
        text: "Next Members ➡️",
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
    { text: "🏢 Platforms", callback_data: `shiftplatforms:${sheetIndex}` },
    { text: "🏠 Dashboard", callback_data: "menu:0" },
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
      `🔄 SHIFTING OVERVIEW
All platforms summary
${sections.length} platform${sections.length === 1 ? "" : "s"} available`,
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

async function getDashboardSheets() {
  const sheets = await listSheetTabs();

  return DASHBOARD_SHEET_TITLES.map((title) => sheets.find((sheet) => sheet.title === title)).filter(
    (sheet): sheet is NonNullable<(typeof sheets)[number]> => Boolean(sheet),
  );
}

async function buildSheetKeyboard() {
  const sheets = await getDashboardSheets();

  return {
    inline_keyboard: sheets.map((sheet, index) => [{
      text: sheet.title === "REAL TIME" ? "📊 REAL TIME BOARD" : "🔄 SHIFTING BOARD",
      callback_data: `sheet:${index}:0`,
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
  const height = 920;

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
              padding: "48px 42px 40px",
              background: "linear-gradient(135deg, #163225 0%, #1f4f46 42%, #1e3a5f 100%)",
              boxShadow: "0 24px 60px rgba(2, 6, 23, 0.35)",
              border: "1px solid rgba(255,255,255,0.10)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "20px",
                  letterSpacing: "3px",
                  textTransform: "uppercase",
                  color: "#bfdbfe",
                  textAlign: "center",
                },
              },
              "Withdraw Team",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "66px",
                  fontWeight: 800,
                  marginTop: "12px",
                  textAlign: "center",
                },
              },
              "Live Operations Center",
            ),
            createElement(
              "div",
              {
                key: "subtitle",
                style: {
                  fontSize: "26px",
                  marginTop: "14px",
                  color: "#e2e8f0",
                  textAlign: "center",
                },
              },
              "Real-time monitoring and shifting control for staff and team leaders.",
            ),
          ],
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
              badge: "📊",
              title: "REAL TIME BOARD",
              subtitle: "Live platform counts, team leaders, daily shift codes, and operational metrics.",
              accent: "linear-gradient(135deg, #0f766e 0%, #1e3a5f 100%)",
            },
            {
              badge: "🔄",
              title: "SHIFTING BOARD",
              subtitle: "Platform overview with Day, Mid, and Night staff assignment control.",
              accent: "linear-gradient(135deg, #1d4ed8 0%, #7c3aed 100%)",
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
  const pageLabel = preview.totalEntryPages > 1 ? `Page ${preview.entryPage + 1} of ${preview.totalEntryPages}` : "Single page";
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
              `${shiftBadge} ${shiftLabel} • ${preview.entries.length} team members • ${pageLabel}`,
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

async function sendMenu(chatId: number, text?: string) {
  const replyMarkup = await buildSheetKeyboard();
  const imageBuffer = await renderDashboardHomeImage();

  await sendTelegramPhoto(
    chatId,
    imageBuffer,
    text ?? "WITHDRAW TEAM
Live Operations Center
Choose one board to open.",
    replyMarkup,
  );
}

async function showSheet(
  callbackQuery: TelegramCallbackQuery,
  sheetIndex: number,
  page: number,
) {
  const message = callbackQuery.message;

  if (!message) {
    await answerCallbackQuery(callbackQuery.id, "The original message is no longer available.");
    return;
  }

  const sheets = await getDashboardSheets();
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

    if (sheet.title === "REAL TIME") {
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

      replyMarkup = buildSectionNavigation(sheetIndex, safePage, REAL_TIME_SECTION_COUNT);
    } else if (sheet.title === "BASICS WITHDARW") {
      const preview = await getBasicsWithdrawPreview();
      imageBuffer = await renderBasicsWithdrawImage(preview);
      caption = buildBasicsCaption();
      replyMarkup = { inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "menu:0" }]] };
    } else if (sheet.title === "WORKFOLIO EMAIL") {
      const preview = await getWorkfolioEmailPreview(page);
      imageBuffer = await renderWorkfolioEmailImage(preview);
      caption = buildWorkfolioEmailCaption(preview);
      replyMarkup = buildSectionNavigation(sheetIndex, preview.page, preview.sections.length);
    } else if (sheet.title === "SHIFTING") {
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
        readSheetWindow(sheet.title, page, config.rowsPerPage, config.columnsToShow),
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
        inline_keyboard: [[{ text: "🏠 Dashboard", callback_data: "menu:0" }]],
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
    await sendMenu(message.chat.id);
    return;
  }

  if (text === "/help") {
    await callTelegram("sendMessage", {
      chat_id: message.chat.id,
      text: `Commands:
/start - open sheet menu
/menu - open sheet menu
/help - show commands
/chatid - show this chat ID

This bot is read-only and now focused on REAL TIME and SHIFTING only.`,
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

    if (message.photo && message.photo.length > 0) {
      await deleteTelegramMessage(chatId, message.message_id);
    }

    await sendMenu(
      chatId,
      "WITHDRAW TEAM Dashboard
Live Operations Center

Choose one board to open.",
    );
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
    const [, sheetIndexValue, pageValue] = data.split(":");
    const sheetIndex = Number(sheetIndexValue);
    const page = Number(pageValue);

    if (Number.isNaN(sheetIndex) || Number.isNaN(page)) {
      await answerCallbackQuery(callbackQuery.id, "Invalid sheet request.");
      return;
    }

    const sheets = await getDashboardSheets();
    const sheet = sheets[sheetIndex];

    if (sheet?.title === "SHIFTING") {
      await showShiftingPlatformMenu(callbackQuery, sheetIndex);
      return;
    }

    await showSheet(callbackQuery, sheetIndex, page);
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

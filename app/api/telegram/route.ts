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
};

type RealTimePreview = {
  timestamp: string;
  headers: [string, string, string, string, string];
  rows: RealTimeRow[];
  hasNextPage: boolean;
  page: number;
  rowOffset: number;
};

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
  nightEntries: ShiftingEntry[];
};

type ShiftingPreview = {
  summary: ShiftingSummaryItem[];
  sections: ShiftingSection[];
  page: number;
  currentSection: ShiftingSection;
};

type SheetNavigation = {
  inline_keyboard: InlineKeyboardButton[][];
};

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

async function getRealTimePreview(page: number, rowsPerPage: number): Promise<RealTimePreview> {
  const safePage = Math.max(0, page);
  const startRow = 3 + safePage * rowsPerPage;
  const endRow = startRow + rowsPerPage;

  const [headerRows, dataRows] = await Promise.all([
    readSheetRange("REAL TIME", "A1:E2"),
    readSheetRange("REAL TIME", `A${startRow}:E${endRow}`),
  ]);

  const timestamp = cleanCell(headerRows[0]?.[0] ?? "REAL TIME");
  const rawHeaders = headerRows[1] ?? ["PLATFORM", "DAY SHIFT", "NIGHT SHIFT", "MID SHIFT", "TOTAL"];
  const headers = [
    normalizeHeader(rawHeaders[0] ?? "PLATFORM"),
    rawHeaders[1] ?? "DAY SHIFT",
    rawHeaders[2] ?? "NIGHT SHIFT",
    rawHeaders[3] ?? "MID SHIFT",
    rawHeaders[4] ?? "TOTAL",
  ] as [string, string, string, string, string];

  const rows = dataRows
    .filter((row) => cleanCell(row[0] ?? "").length > 0)
    .map((row) => ({
      platform: cleanCell(row[0] ?? "-"),
      dayShift: cleanCell(row[1] ?? "-"),
      nightShift: cleanCell(row[2] ?? "-"),
      midShift: cleanCell(row[3] ?? "-"),
      total: cleanCell(row[4] ?? "-"),
    }));

  return {
    timestamp,
    headers,
    rows: rows.slice(0, rowsPerPage),
    hasNextPage: rows.length > rowsPerPage,
    page: safePage,
    rowOffset: startRow,
  };
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

async function getShiftingPreview(page: number): Promise<ShiftingPreview> {
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
      currentSection.dayEntries.push(dayEntry);
    }

    if (nightEntry && nightEntry.name) {
      currentSection.nightEntries.push(nightEntry);
    }
  }

  const safePage = Math.max(0, Math.min(page, Math.max(sections.length - 1, 0)));
  const fallbackSection: ShiftingSection = {
    title: "SHIFTING",
    dayEntries: [],
    nightEntries: [],
  };

  return {
    summary,
    sections,
    page: safePage,
    currentSection: sections[safePage] ?? fallbackSection,
  };
}

function buildSheetCaption(sheetTitle: string, rowOffset: number, rowCount: number, rows: string[][]) {
  const lastVisibleRow = rows.length > 0 ? rowOffset + rows.length - 1 : rowOffset;
  return `${sheetTitle}\nRows ${rowOffset}-${lastVisibleRow} of ${rowCount}`;
}

function buildRealTimeCaption(preview: RealTimePreview) {
  const lastVisibleRow = preview.rows.length > 0 ? preview.rowOffset + preview.rows.length - 1 : preview.rowOffset;
  return `REAL TIME\n${preview.timestamp}\nPlatforms ${preview.rowOffset - 2}-${lastVisibleRow - 2}`;
}

function buildShiftingCaption(preview: ShiftingPreview) {
  return `SHIFTING\n${preview.currentSection.title}\nSection ${preview.page + 1} of ${Math.max(preview.sections.length, 1)}`;
}

function buildSheetNavigation(sheetIndex: number, page: number, hasNextPage: boolean): SheetNavigation {
  const inlineKeyboard: InlineKeyboardButton[][] = [];
  const navigationButtons: InlineKeyboardButton[] = [];

  if (page > 0) {
    navigationButtons.push({
      text: "Previous",
      callback_data: `sheet:${sheetIndex}:${page - 1}`,
    });
  }

  if (hasNextPage) {
    navigationButtons.push({
      text: "Next",
      callback_data: `sheet:${sheetIndex}:${page + 1}`,
    });
  }

  if (navigationButtons.length > 0) {
    inlineKeyboard.push(navigationButtons);
  }

  inlineKeyboard.push([{ text: "All sheets", callback_data: "menu:0" }]);

  return { inline_keyboard: inlineKeyboard };
}

function buildSectionNavigation(sheetIndex: number, page: number, totalSections: number): SheetNavigation {
  return buildSheetNavigation(sheetIndex, page, page + 1 < totalSections);
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

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Telegram API ${method} failed: ${response.status} ${body}`);
  }
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

async function buildSheetKeyboard() {
  const sheets = await listSheetTabs();
  const buttons: InlineKeyboardButton[][] = [];

  for (let index = 0; index < sheets.length; index += 2) {
    buttons.push(
      sheets.slice(index, index + 2).map((sheet, offset) => ({
        text: sheet.title,
        callback_data: `sheet:${index + offset}:0`,
      })),
    );
  }

  return {
    inline_keyboard: buttons,
  };
}

async function renderGenericSheetImage(
  sheetTitle: string,
  rowOffset: number,
  rowCount: number,
  rows: string[][],
) {
  const displayRows = buildDisplayRows(rows, rowOffset);
  const width = 1200;
  const height = Math.max(720, 240 + displayRows.length * 88);
  const lastVisibleRow = rows.length > 0 ? rowOffset + rows.length - 1 : rowOffset;

  const image = new ImageResponse(
    createElement(
      "div",
      {
        style: {
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          background: "linear-gradient(180deg, #101726 0%, #1a2436 100%)",
          color: "#f8fafc",
          padding: "42px",
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
              marginBottom: "26px",
              padding: "28px 32px",
              borderRadius: "28px",
              background: "rgba(15, 23, 42, 0.82)",
              border: "1px solid rgba(148, 163, 184, 0.18)",
            },
          },
          [
            createElement(
              "div",
              {
                key: "eyebrow",
                style: {
                  fontSize: "24px",
                  letterSpacing: "2px",
                  textTransform: "uppercase",
                  color: "#f59e0b",
                  marginBottom: "12px",
                },
              },
              "Telegram Sheet View",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "48px",
                  fontWeight: 700,
                  marginBottom: "10px",
                },
              },
              sheetTitle,
            ),
            createElement(
              "div",
              {
                key: "meta",
                style: {
                  fontSize: "24px",
                  color: "#cbd5e1",
                },
              },
              `Rows ${rowOffset}-${lastVisibleRow} of ${rowCount}`,
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "rows",
            style: {
              display: "flex",
              flexDirection: "column",
              gap: "14px",
            },
          },
          displayRows.map((row, index) =>
            createElement(
              "div",
              {
                key: `${row.rowLabel}-${index}`,
                style: {
                  display: "flex",
                  alignItems: "center",
                  padding: "18px 22px",
                  borderRadius: "22px",
                  background: index % 2 === 0 ? "rgba(30, 41, 59, 0.92)" : "rgba(15, 23, 42, 0.92)",
                  border: "1px solid rgba(148, 163, 184, 0.14)",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: `label-${row.rowLabel}`,
                    style: {
                      minWidth: "74px",
                      padding: "10px 14px",
                      borderRadius: "14px",
                      background: "rgba(245, 158, 11, 0.16)",
                      color: "#fcd34d",
                      fontSize: "26px",
                      fontWeight: 700,
                      textAlign: "center",
                      marginRight: "18px",
                    },
                  },
                  row.rowLabel,
                ),
                createElement(
                  "div",
                  {
                    key: `content-${row.rowLabel}`,
                    style: {
                      display: "flex",
                      flex: 1,
                      fontSize: "28px",
                      color: "#e2e8f0",
                    },
                  },
                  row.content,
                ),
              ],
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

async function renderRealTimeImage(preview: RealTimePreview) {
  const width = 1200;
  const height = Math.max(860, 300 + preview.rows.length * 72);
  const columns = [
    { key: "platform", label: preview.headers[0], width: 350 },
    { key: "dayShift", label: preview.headers[1], width: 180 },
    { key: "nightShift", label: preview.headers[2], width: 180 },
    { key: "midShift", label: preview.headers[3], width: 180 },
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
          background: "#f4f1ea",
          color: "#111827",
          padding: "36px",
          fontFamily: "Georgia, serif",
          boxSizing: "border-box",
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
              background: "#3f725c",
              color: "#ffffff",
              borderRadius: "20px 20px 0 0",
              padding: "22px 28px",
              border: "3px solid #244536",
              borderBottom: "0",
              alignItems: "center",
            },
          },
          [
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "28px",
                  letterSpacing: "2px",
                  marginBottom: "8px",
                  textTransform: "uppercase",
                },
              },
              "REAL TIME",
            ),
            createElement(
              "div",
              {
                key: "time",
                style: {
                  fontSize: "54px",
                  fontWeight: 700,
                },
              },
              preview.timestamp,
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "table",
            style: {
              display: "flex",
              flexDirection: "column",
              border: "3px solid #244536",
              borderTop: "0",
            },
          },
          [
            createElement(
              "div",
              {
                key: "thead",
                style: {
                  display: "flex",
                  background: "#3f725c",
                  color: "#ffffff",
                  borderBottom: "3px solid #244536",
                },
              },
              columns.map((column) =>
                createElement(
                  "div",
                  {
                    key: column.key,
                    style: {
                      width: `${column.width}px`,
                      padding: "16px 14px",
                      borderRight: column.key === "total" ? "0" : "2px solid #244536",
                      textAlign: "center",
                      fontSize: "24px",
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
                    background: index % 2 === 0 ? "#ffffff" : "#f6f7f8",
                    borderBottom:
                      index === preview.rows.length - 1 ? "0" : "2px solid #1f2937",
                  },
                },
                columns.map((column) =>
                  createElement(
                    "div",
                    {
                      key: `${row.platform}-${column.key}`,
                      style: {
                        width: `${column.width}px`,
                        padding: "14px 12px",
                        borderRight: column.key === "total" ? "0" : "2px solid #1f2937",
                        textAlign: column.key === "platform" ? "left" : "center",
                        fontSize: column.key === "platform" ? "24px" : "28px",
                        fontWeight: column.key === "platform" ? 700 : 600,
                      },
                    },
                    row[column.key],
                  ),
                ),
              ),
            ),
          ],
        ),
        createElement(
          "div",
          {
            key: "footer",
            style: {
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginTop: "18px",
              color: "#334155",
              fontSize: "22px",
            },
          },
          [
            createElement(
              "div",
              { key: "range" },
              `Platforms ${preview.rowOffset - 2}-${preview.rowOffset - 2 + Math.max(preview.rows.length - 1, 0)}`,
            ),
            createElement(
              "div",
              { key: "page" },
              `Page ${preview.page + 1}`,
            ),
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

function renderShiftingEntryCard(entry: ShiftingEntry, accent: string) {
  return createElement(
    "div",
    {
      key: `${entry.name}-${entry.id}-${entry.shift}`,
      style: {
        display: "flex",
        flexDirection: "column",
        gap: "6px",
        padding: "14px 16px",
        borderRadius: "18px",
        background: "rgba(255,255,255,0.92)",
        border: `2px solid ${accent}`,
        marginBottom: "12px",
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
            gap: "12px",
            fontSize: "18px",
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
            fontSize: "24px",
            fontWeight: 700,
            color: "#0f172a",
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
          },
        },
        shortenCell(entry.account || entry.startDate || "", 44),
      ),
    ],
  );
}

async function renderShiftingImage(preview: ShiftingPreview) {
  const width = 1280;
  const maxEntries = Math.max(preview.currentSection.dayEntries.length, preview.currentSection.nightEntries.length, 1);
  const summaryRows = Math.ceil(Math.max(preview.summary.length, 1) / 4);
  const height = Math.max(900, 300 + summaryRows * 110 + maxEntries * 132);
  const summaryAccent = ["#1d4ed8", "#0f766e", "#9333ea", "#c2410c"];

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
              background: "#1f4f46",
              color: "#ffffff",
              borderRadius: "24px",
              padding: "24px 28px",
              marginBottom: "20px",
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
                },
              },
              "Shifting Overview",
            ),
            createElement(
              "div",
              {
                key: "title",
                style: {
                  fontSize: "46px",
                  fontWeight: 800,
                  marginTop: "8px",
                },
              },
              preview.currentSection.title,
            ),
            createElement(
              "div",
              {
                key: "page",
                style: {
                  fontSize: "22px",
                  color: "#dbeafe",
                  marginTop: "8px",
                },
              },
              `Section ${preview.page + 1} of ${Math.max(preview.sections.length, 1)}`,
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
              gap: "14px",
              marginBottom: "20px",
            },
          },
          preview.summary.map((item, index) =>
            createElement(
              "div",
              {
                key: item.label,
                style: {
                  width: "290px",
                  display: "flex",
                  flexDirection: "column",
                  padding: "16px 18px",
                  borderRadius: "18px",
                  background: "#ffffff",
                  borderTop: `6px solid ${summaryAccent[index % summaryAccent.length]}`,
                  boxShadow: "0 8px 24px rgba(15, 23, 42, 0.08)",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "label",
                    style: {
                      fontSize: "18px",
                      color: "#475569",
                    },
                  },
                  item.label,
                ),
                createElement(
                  "div",
                  {
                    key: "value",
                    style: {
                      fontSize: "38px",
                      fontWeight: 800,
                      color: "#0f172a",
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
            key: "columns",
            style: {
              display: "flex",
              gap: "18px",
              flex: 1,
            },
          },
          [
            createElement(
              "div",
              {
                key: "day",
                style: {
                  display: "flex",
                  flexDirection: "column",
                  flex: 1,
                  background: "rgba(255,255,255,0.72)",
                  borderRadius: "24px",
                  padding: "18px",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "title",
                    style: {
                      fontSize: "28px",
                      fontWeight: 800,
                      color: "#0f766e",
                      marginBottom: "12px",
                    },
                  },
                  `Day / Mid (${preview.currentSection.dayEntries.length})`,
                ),
                ...(preview.currentSection.dayEntries.length > 0
                  ? preview.currentSection.dayEntries.map((entry) =>
                      renderShiftingEntryCard(entry, "#0f766e"),
                    )
                  : [
                      createElement(
                        "div",
                        {
                          key: "empty-day",
                          style: {
                            fontSize: "22px",
                            color: "#64748b",
                            padding: "18px 8px",
                          },
                        },
                        "No day entries in this section.",
                      ),
                    ]),
              ],
            ),
            createElement(
              "div",
              {
                key: "night",
                style: {
                  display: "flex",
                  flexDirection: "column",
                  flex: 1,
                  background: "rgba(255,255,255,0.72)",
                  borderRadius: "24px",
                  padding: "18px",
                },
              },
              [
                createElement(
                  "div",
                  {
                    key: "title",
                    style: {
                      fontSize: "28px",
                      fontWeight: 800,
                      color: "#1d4ed8",
                      marginBottom: "12px",
                    },
                  },
                  `Night (${preview.currentSection.nightEntries.length})`,
                ),
                ...(preview.currentSection.nightEntries.length > 0
                  ? preview.currentSection.nightEntries.map((entry) =>
                      renderShiftingEntryCard(entry, "#1d4ed8"),
                    )
                  : [
                      createElement(
                        "div",
                        {
                          key: "empty-night",
                          style: {
                            fontSize: "22px",
                            color: "#64748b",
                            padding: "18px 8px",
                          },
                        },
                        "No night entries in this section.",
                      ),
                    ]),
              ],
            ),
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

  await callTelegram("sendMessage", {
    chat_id: chatId,
    text:
      text ??
      "Choose a sheet below. The bot reads your private spreadsheet and shows it here in Telegram.",
    reply_markup: replyMarkup,
  });
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

  const sheets = await listSheetTabs();
  const sheet = sheets[sheetIndex];

  if (!sheet) {
    await answerCallbackQuery(callbackQuery.id, "That sheet was not found.");
    return;
  }

  let imageBuffer: ArrayBuffer;
  let caption: string;
  let replyMarkup: SheetNavigation;

  if (sheet.title === "REAL TIME") {
    const preview = await getRealTimePreview(page, getPreviewRows());
    imageBuffer = await renderRealTimeImage(preview);
    caption = buildRealTimeCaption(preview);
    replyMarkup = buildSheetNavigation(sheetIndex, page, preview.hasNextPage);
  } else if (sheet.title === "SHIFTING") {
    const preview = await getShiftingPreview(page);
    imageBuffer = await renderShiftingImage(preview);
    caption = buildShiftingCaption(preview);
    replyMarkup = buildSectionNavigation(sheetIndex, preview.page, preview.sections.length);
  } else {
    const rowsPerPage = getPreviewRows();
    const columnsToShow = getPreviewColumns();
    const window = await readSheetWindow(sheet.title, page, rowsPerPage, columnsToShow);
    imageBuffer = await renderGenericSheetImage(
      sheet.title,
      window.rowOffset,
      sheet.rowCount,
      window.rows,
    );
    caption = buildSheetCaption(sheet.title, window.rowOffset, sheet.rowCount, window.rows);
    replyMarkup = buildSheetNavigation(sheetIndex, page, window.hasNextPage);
  }

  await answerCallbackQuery(callbackQuery.id);

  if (message.photo && message.photo.length > 0) {
    await deleteTelegramMessage(message.chat.id, message.message_id);
  }

  await sendTelegramPhoto(message.chat.id, imageBuffer, caption, replyMarkup);
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

  if (data.startsWith("menu:")) {
    await answerCallbackQuery(callbackQuery.id);

    if (message.photo && message.photo.length > 0) {
      await deleteTelegramMessage(chatId, message.message_id);
    }

    await sendMenu(
      chatId,
      "Choose a sheet below. The bot reads your private spreadsheet and shows it here in Telegram.",
    );
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

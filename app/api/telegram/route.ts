import { ImageResponse } from "next/og";
import { createElement } from "react";
import { NextResponse } from "next/server";
import { listSheetTabs, readSheetWindow } from "@/lib/googleSheets";

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

  return `${value.slice(0, maxLength - 1)}…`;
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

function buildSheetCaption(sheetTitle: string, rowOffset: number, rowCount: number, rows: string[][]) {
  const lastVisibleRow = rows.length > 0 ? rowOffset + rows.length - 1 : rowOffset;
  return `${sheetTitle}
Rows ${rowOffset}-${lastVisibleRow} of ${rowCount}`;
}

function buildSheetNavigation(sheetIndex: number, page: number, hasNextPage: boolean) {
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
  replyMarkup: { inline_keyboard: InlineKeyboardButton[][] },
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

async function renderSheetImage(
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

  const rowsPerPage = getPreviewRows();
  const columnsToShow = getPreviewColumns();
  const window = await readSheetWindow(sheet.title, page, rowsPerPage, columnsToShow);
  const imageBuffer = await renderSheetImage(
    sheet.title,
    window.rowOffset,
    sheet.rowCount,
    window.rows,
  );
  const caption = buildSheetCaption(
    sheet.title,
    window.rowOffset,
    sheet.rowCount,
    window.rows,
  );
  const replyMarkup = buildSheetNavigation(sheetIndex, page, window.hasNextPage);

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

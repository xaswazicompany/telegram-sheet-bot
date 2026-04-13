import { NextResponse } from "next/server";
import { listSheetTabs, readSheetWindow } from "@/lib/googleSheets";

type TelegramChat = {
  id: number;
};

type TelegramMessage = {
  message_id: number;
  chat: TelegramChat;
  text?: string;
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

function escapeHtml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function isAllowedChat(chatId: number) {
  const allowedChatIds = getAllowedChatIds();

  if (allowedChatIds.length === 0) {
    return true;
  }

  return allowedChatIds.includes(String(chatId));
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

async function answerCallbackQuery(callbackQueryId: string, text?: string) {
  await callTelegram("answerCallbackQuery", {
    callback_query_id: callbackQueryId,
    text,
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

function formatRows(rows: string[][], rowOffset: number) {
  const lines = rows
    .filter((row) => row.some((cell) => cell.length > 0))
    .map((row, index) => {
      const values = row.map((cell) => (cell.length > 0 ? cell : "-"));
      return `${rowOffset + index}. ${values.join(" | ")}`;
    });

  if (lines.length === 0) {
    return "No data found in this range yet.";
  }

  return lines.join("\n");
}

function truncateTelegramHtml(text: string, maxLength = 3900) {
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, maxLength - 20)}\n...truncated`;
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
  const lastVisibleRow =
    window.rows.length > 0 ? window.rowOffset + window.rows.length - 1 : window.rowOffset;
  const text = truncateTelegramHtml(
    [
      `<b>${escapeHtml(sheet.title)}</b>`,
      `Rows ${window.rowOffset}-${lastVisibleRow} of ${sheet.rowCount}`,
      `Showing up to ${columnsToShow} columns per row.`,
      "",
      `<pre>${escapeHtml(formatRows(window.rows, window.rowOffset))}</pre>`,
    ].join("\n"),
  );

  const navigationButtons: InlineKeyboardButton[] = [];

  if (page > 0) {
    navigationButtons.push({
      text: "Previous",
      callback_data: `sheet:${sheetIndex}:${page - 1}`,
    });
  }

  if (window.hasNextPage) {
    navigationButtons.push({
      text: "Next",
      callback_data: `sheet:${sheetIndex}:${page + 1}`,
    });
  }

  const inlineKeyboard: InlineKeyboardButton[][] = [];

  if (navigationButtons.length > 0) {
    inlineKeyboard.push(navigationButtons);
  }

  inlineKeyboard.push([{ text: "All sheets", callback_data: "menu:0" }]);

  await answerCallbackQuery(callbackQuery.id);
  await callTelegram("editMessageText", {
    chat_id: message.chat.id,
    message_id: message.message_id,
    text,
    parse_mode: "HTML",
    reply_markup: {
      inline_keyboard: inlineKeyboard,
    },
  });
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
      text: `This chat ID is: ${message.chat.id}` ,
    });
    return;
  }

  await callTelegram("sendMessage", {
    chat_id: message.chat.id,
    text: "Use /start to open the sheet menu.",
  });
}

async function handleCallbackQuery(callbackQuery: TelegramCallbackQuery) {
  const chatId = callbackQuery.message?.chat.id;

  if (!chatId) {
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
    const replyMarkup = await buildSheetKeyboard();

    await callTelegram("editMessageText", {
      chat_id: chatId,
      message_id: callbackQuery.message?.message_id,
      text:
        "Choose a sheet below. The bot reads your private spreadsheet and shows it here in Telegram.",
      reply_markup: replyMarkup,
    });
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

# Professional Website + Google Sheets Starter

This project gives you a professional website starter that sends form submissions to Google Sheets using a secure backend. It now also includes a private Telegram viewer so staff can check spreadsheet tabs without direct access to the Google Sheet or Excel file.

## Stack

- Next.js with TypeScript
- Server-side API route for form submission
- Google Sheets API with a Google service account
- Telegram Bot webhook for private sheet viewing

## Why this is professional

The browser never talks to Google Sheets directly. Instead, your website sends data to your own backend route at `/api/leads`, and the server writes to Google Sheets using private credentials stored in environment variables. Telegram users only see the bot output, not the sheet itself.

## 1. Install dependencies

```bash
npm install
```

## 2. Create your Google Sheet

Create a spreadsheet and add this first row in your sheet:

```text
Timestamp | Full Name | Email | Company | Project Type | Budget | Message
```

Rename the sheet tab to `Leads`, or change `GOOGLE_SHEETS_SHEET_NAME` in your environment file.

## 3. Create a Google Cloud project

1. Open Google Cloud Console.
2. Create a new project.
3. Enable the Google Sheets API.
4. Create a service account.
5. Create a JSON key for that service account.
6. Copy the service account email and private key into `.env.local`.

## 4. Share the sheet with the service account

Open your Google Sheet and share it with the service account email, like you would share it with a normal person. Give it Editor access.

## 5. Add environment variables

Create `.env.local`:

```bash
cp .env.example .env.local
```

Then fill in the real values:

- `GOOGLE_CLIENT_EMAIL`
- `GOOGLE_PRIVATE_KEY`
- `GOOGLE_SHEETS_SPREADSHEET_ID`
- `GOOGLE_SHEETS_SHEET_NAME`
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_ALLOWED_CHAT_IDS`
- `TELEGRAM_SHEET_PREVIEW_ROWS`
- `TELEGRAM_SHEET_PREVIEW_COLUMNS`

Important:
Keep the private key wrapped in quotes and preserve the `\n` line breaks.

## 6. Run the website

```bash
npm run dev
```

Open `http://localhost:3000`

## How it works

- The form in `app/page.tsx` sends data to `POST /api/leads`
- The API route in `app/api/leads/route.ts` validates the request
- The Google Sheets helper in `lib/googleSheets.ts` appends a new row to your sheet
- The Telegram webhook in `app/api/telegram/route.ts` lists all spreadsheet tabs and shows read-only previews inside Telegram

## Telegram bot viewer

Use this when you want staff, TLs, or managers to check the spreadsheet in Telegram without giving them access to the spreadsheet itself.

### What the bot does

- `/start` or `/menu` shows all sheet tabs as Telegram buttons
- Tapping a button opens a read-only preview of that sheet
- `Previous` and `Next` buttons let users page through rows
- Access can be restricted to approved chat IDs only

### Environment variables

Add these to `.env.local`:

```bash
TELEGRAM_BOT_TOKEN=123456789:your_bot_token
TELEGRAM_ALLOWED_CHAT_IDS=123456789,-1001234567890
TELEGRAM_SHEET_PREVIEW_ROWS=10
TELEGRAM_SHEET_PREVIEW_COLUMNS=8
```

Notes:

- `TELEGRAM_ALLOWED_CHAT_IDS` is optional but recommended. Use a comma-separated list of private chat IDs or group IDs.
- Group IDs often start with `-100`.
- If `TELEGRAM_ALLOWED_CHAT_IDS` is empty, any Telegram chat that can reach the bot will be allowed.

### Connect your existing Telegram bot

1. Create or reuse your bot with BotFather.
2. Put the real bot token in `.env.local`.
3. Deploy this app to a public HTTPS URL.
4. Set the Telegram webhook to your deployed route: `https://your-domain.com/api/telegram`
5. In Telegram, send `/start` to the bot.

To register the webhook, call Telegram's `setWebhook` for your bot token and point it to the URL above.

### Privacy model

- Users never get the Google Sheets link
- Users never get the service account credentials
- The bot only returns the rows shown in Telegram
- You control who can open the bot by chat ID

## Deploy

You can deploy this to Vercel or another Node-compatible host.

Before deploying, add the same environment variables in your hosting dashboard.

## Next improvements

- Add spam protection with reCAPTCHA or Cloudflare Turnstile
- Add a success email notification
- Add an admin dashboard
- Save multiple forms to different tabs in the same spreadsheet
- Add sheet-specific formatting for key tabs such as `REAL TIME` or `SHIFTING`

# Excel Bill Maker — Telegram Bot

Use `/new` in Telegram to open a simple invoice form. The bot asks for the
invoice and buyer details, then item rows, then GST. Seller details remain fixed
in the supplied invoice template. After the final field, it sends an editable
Excel invoice and its PDF.

## Setup on Windows

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Copy `.env.example` to `.env` and put the Telegram bot token there:

```text
TELEGRAM_BOT_TOKEN=your-token-here
```

Start the bot:

```powershell
.\.venv\Scripts\python.exe main.py
```

## Create an invoice

- `/new` or `/start` starts a new invoice form.
- Enter `AUTO` for the next invoice number or provide your own number.
- Enter `-` for optional fields you want to leave blank.
- Enter each item's description, HSN/SAC, quantity in Qtl, and rate per Qtl.
  For reference, 12 tonnes = 120 Qtl; put a count such as 300 katta in the
  description, e.g. `Mungfali (300 katta)`.
- Choose IGST or CGST+SGST and enter the total GST rate.
- The bot sends `invoice_<number>.xlsx` and `invoice_<number>.pdf`. The PDF
  renderer reads the generated workbook's cell values, merges, widths, and
  styles and lays it out on A4 using Python packages. No Microsoft Office or
  LibreOffice installation is needed.

The form includes invoice date, buyer name/address/GSTIN, buyer order number,
dispatch document number, delivery note, payment terms, vehicle number,
destination, and delivery terms. The template supports up to 11 item rows.

Seller details are not collected in chat. The Excel seller information stays
fixed in `Hanuman.xlsx`; both the Excel and its PDF therefore use the same seller
details and invoice layout.

## Files

| File | Purpose |
|---|---|
| `main.py` | Telegram form and file delivery |
| `api/webhook.py` | Vercel Telegram webhook endpoint |
| `upstash_persistence.py` | Saves form progress across serverless requests |
| `set_webhook.py` | Registers the deployed webhook with Telegram |
| `vercel.json` | Vercel function configuration |
| `billing.py` | GST math and amount-in-words conversion |
| `excel_generator.py` | Fills a copy of the existing Excel template |
| `pdf_converter.py` | Renders the generated Excel workbook to PDF using Python packages |
| `requirements.txt` | Python dependencies |
| `.python-version` | Pins the Vercel build to Python 3.12 |
| `.env.example` | Environment-variable template; contains no live credentials |
| `invoice_counter.json` | Auto-incrementing invoice number state |
| `generated_invoices/` | Saved Excel and PDF invoices |

## Vercel webhook deployment

Vercel can run the Python webhook without Excel or LibreOffice. Each Telegram
update starts a serverless request, so the conversation state and invoice
number counter use an Upstash Redis database instead of local files.

1. Push this project (including `Hanuman.xlsx`) to a private GitHub repository.
   Do not commit `.env` or paste its secrets into chat.
2. Create an Upstash Redis database and copy its REST URL and token.
3. Import the repository into Vercel. Add these Project Environment Variables
   for Production and Preview:

   ```text
   TELEGRAM_BOT_TOKEN=your-new-bot-token
   TELEGRAM_WEBHOOK_SECRET=a-long-random-value-using-letters-numbers-_-
   UPSTASH_REDIS_REST_URL=https://your-database.upstash.io
   UPSTASH_REDIS_REST_TOKEN=your-upstash-token
   ```

4. Deploy. In a local copy, put the same bot token and webhook secret in `.env`,
   install requirements, then register the webhook once:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   .\.venv\Scripts\python.exe set_webhook.py https://your-project.vercel.app
   ```

5. Send `/new` to the bot. Do not run local polling with `main.py` at the same
   time as the Vercel webhook; Telegram supports one update delivery mode per
   bot, and polling will replace the webhook.

The Python runtime on Vercel is currently Beta. Generated files are written to
temporary function storage and sent to Telegram; they are not a permanent
archive. Keep the original GitHub repository and a backup of the workbook
template. If Telegram webhook registration is changed or the deployment URL
changes, run `set_webhook.py` again. The Hobby plan is restricted to personal,
non-commercial use, so it is not eligible for a real business invoice bot;
commercial use requires Vercel Pro or Enterprise. Upstash credentials and
usage limits are governed by the selected Upstash plan.

## Known constraints

- The template has room for at most 11 invoice items.
- The PDF keeps workbook font sizes, bold weight, alignments, colors, and
  proportions while scaling the sheet to one A4 page. Arial is rendered with
  the built-in Helvetica-compatible PDF font. Amounts use `Rs.` so no Office
  font installation is needed.
- The PDF renderer uses the generated workbook's template but replaces its
  formula cells in a temporary copy because Python does not calculate Excel
  formulas. The editable `.xlsx` remains formula-based.
- A failed Vercel build or missing environment variable prevents the webhook
  from responding; check Vercel Function Logs and confirm all four variables.

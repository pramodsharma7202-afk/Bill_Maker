"""
Telegram GST Bill Maker Bot
===========================
The Vercel webhook opens a stateless Telegram Mini App form and sends back an
editable Excel invoice plus a PDF styled after Hanuman.xlsx. Local polling
continues to support the chat-based wizard.

SETUP
-----
1. pip install -r requirements.txt
2. Set your bot token as an environment variable (do NOT hardcode it):
      export TELEGRAM_BOT_TOKEN="123456:ABC-your-token"
   (On Windows PowerShell:  $env:TELEGRAM_BOT_TOKEN="123456:ABC-your-token")
3. python main.py

USAGE
-----
In Telegram, message your bot:
   /start or /new   -> open a simple invoice input form
   /cancel          -> discard the current invoice form
"""

import asyncio
import json
import logging
import math
import os
import re
from datetime import date

from dotenv import load_dotenv
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
    WebAppInfo,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

from billing import calculate_gst
from excel_generator import generate_invoice_xlsx
from pdf_converter import convert_xlsx_to_pdf

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

# ---- Conversation states ---------------------------------------------------
(
    INVOICE_NO,
    BUYER_NAME,
    BUYER_ADDRESS,
    BUYER_GSTIN,
    INVOICE_DATE,
    DELIVERY_NOTE,
    PAYMENT_TERMS,
    VEHICLE_NO,
    DESTINATION,
    BUYER_ORDER_NO,
    DISPATCH_DOC_NO,
    DELIVERY_TERMS,
    ITEM_DESC,
    ITEM_HSN,
    ITEM_QTY,
    ITEM_RATE,
    ADD_MORE,
    GST_TYPE,
    GST_RATE,
) = range(19)

COUNTER_FILE = os.path.join(os.path.dirname(__file__), "invoice_counter.json")
INVOICE_OUTPUT_DIR = os.environ.get("INVOICE_OUTPUT_DIR") or (
    "/tmp/generated_invoices" if os.environ.get("VERCEL") else
    os.path.join(os.path.dirname(__file__), "generated_invoices")
)


def _template_invoice_no() -> int:
    try:
        from openpyxl import load_workbook

        template_path = os.path.join(os.path.dirname(__file__), "Hanuman.xlsx")
        return int(load_workbook(template_path, read_only=True, data_only=True)["Tax Invoice"]["G4"].value or 0)
    except (OSError, ValueError, KeyError, TypeError):
        return 0


async def _next_invoice_no() -> int:
    """Increment the local polling bot's counter past the template number."""
    data = {"last": 0}
    if os.path.exists(COUNTER_FILE):
        try:
            with open(COUNTER_FILE) as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    template_number = _template_invoice_no()
    candidate = max(int(data.get("last", 0)), template_number) + 1
    output_dir = INVOICE_OUTPUT_DIR
    while any(os.path.exists(os.path.join(output_dir, f"invoice_{candidate}.{ext}")) for ext in ("xlsx", "pdf")):
        candidate += 1
    data["last"] = candidate
    with open(COUNTER_FILE, "w") as f:
        json.dump(data, f)
    return data["last"]


async def _record_manual_invoice_no(value: str) -> None:
    """Keep future automatic numbers beyond a manually entered numeric number."""
    try:
        number = int(value)
    except ValueError:
        return
    data = {"last": 0}
    if os.path.exists(COUNTER_FILE):
        try:
            with open(COUNTER_FILE) as counter:
                data = json.load(counter)
        except (json.JSONDecodeError, OSError):
            pass
    data["last"] = max(int(data.get("last", 0)), number)
    with open(COUNTER_FILE, "w") as counter:
        json.dump(data, counter)


async def _reserve_invoice_no(value: str) -> bool:
    stem = f"invoice_{value}"
    return not any(os.path.exists(os.path.join(INVOICE_OUTPUT_DIR, f"{stem}.{ext}")) for ext in ("xlsx", "pdf"))


def _invoice_form_url() -> str | None:
    """Return the stable Vercel URL used by Telegram's Mini App button."""
    domain = (
        os.environ.get("VERCEL_PROJECT_PRODUCTION_URL")
        or os.environ.get("VERCEL_URL")
        or os.environ.get("INVOICE_APP_URL")
    )
    if not domain:
        return None
    base_url = domain if domain.startswith("https://") else f"https://{domain}"
    return f"{base_url.rstrip('/')}/invoice-form.html"


def _parse_float(text: str) -> float:
    return float(text.replace(",", "").strip())


async def _deliver_invoice(message, invoice_data: dict) -> None:
    out_dir = INVOICE_OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    stem = f"invoice_{invoice_data['invoice_no']}"
    xlsx_path = os.path.join(out_dir, f"{stem}.xlsx")
    pdf_path = os.path.join(out_dir, f"{stem}.pdf")
    if os.path.exists(xlsx_path) or os.path.exists(pdf_path):
        raise FileExistsError(f"Invoice number {invoice_data['invoice_no']} already has generated files.")
    generate_invoice_xlsx(invoice_data, xlsx_path)
    convert_xlsx_to_pdf(xlsx_path, pdf_path, invoice_data)
    gst = invoice_data["gst"]
    caption = f"Invoice No. {invoice_data['invoice_no']}\nGrand Total: Rs. {gst['grand_total']:,.2f}"
    for file_path, label in ((xlsx_path, "Editable Excel invoice"), (pdf_path, "PDF invoice")):
        with open(file_path, "rb") as invoice_file:
            await message.reply_document(
                document=invoice_file,
                filename=os.path.basename(file_path),
                caption=f"{label}\n{caption}",
            )


# ---- Handlers ---------------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    context.user_data["items"] = []
    form_url = _invoice_form_url()
    if os.environ.get("VERCEL") and not form_url:
        await update.message.reply_text(
            "Invoice form URL is unavailable. Enable Vercel system environment "
            "variables or set INVOICE_APP_URL to the production domain."
        )
        return ConversationHandler.END
    if form_url:
        keyboard = ReplyKeyboardMarkup(
            [[KeyboardButton("🧾 Open GST Invoice Form", web_app=WebAppInfo(url=form_url))]],
            resize_keyboard=True,
            is_persistent=True,
        )
        await update.message.reply_text(
            "🧾 Tap the button below, fill the invoice form, and press Generate. "
            "Your unfinished draft autosaves in Telegram CloudStorage.",
            reply_markup=keyboard,
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "🧾 New GST invoice form\n\n"
        "Invoice number? Send a number to use your own, or AUTO for the next number.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return INVOICE_NO


def _form_text(payload: dict, key: str, label: str, *, required=False, limit=250):
    value = str(payload.get(key, "") or "").strip()
    if required and not value:
        raise ValueError(f"{label} is required.")
    if len(value) > limit:
        raise ValueError(f"{label} is too long (maximum {limit} characters).")
    return value or None


def _invoice_from_web_app(payload: dict, update: Update) -> dict:
    if not isinstance(payload, dict):
        raise ValueError("The form data is not valid.")

    raw_date = _form_text(payload, "invoice_date", "Invoice date", required=True, limit=10)
    try:
        invoice_date_value = date.fromisoformat(raw_date)
    except ValueError as exc:
        raise ValueError("Please select a valid invoice date.") from exc

    raw_number = _form_text(payload, "invoice_no", "Invoice number", limit=24) or "AUTO"
    if raw_number.casefold() in {"auto", "-", "blank"}:
        # Telegram update IDs are unique for a bot, allowing no-database
        # invoices to receive a unique identifier without a shared counter.
        invoice_no = f"INV-{invoice_date_value:%y%m%d}-{update.update_id}"
    elif re.fullmatch(r"[A-Za-z0-9_-]{1,24}", raw_number):
        invoice_no = raw_number
    else:
        raise ValueError("Invoice number can use letters, numbers, _ and - only.")

    invoice_date = invoice_date_value.strftime("%d-%m-%Y")

    item_payloads = payload.get("items")
    if not isinstance(item_payloads, list) or not 1 <= len(item_payloads) <= 11:
        raise ValueError("Add between 1 and 11 items.")
    items = []
    for index, raw_item in enumerate(item_payloads, start=1):
        if not isinstance(raw_item, dict):
            raise ValueError(f"Item {index} is invalid.")
        description = _form_text(raw_item, "description", f"Item {index} description", required=True, limit=180)
        hsn = _form_text(raw_item, "hsn", f"Item {index} HSN/SAC", required=True, limit=32)
        try:
            qty = _parse_float(str(raw_item.get("qty", "")))
            rate = _parse_float(str(raw_item.get("rate", "")))
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Item {index} quantity and rate must be numbers.") from exc
        if not (math.isfinite(qty) and math.isfinite(rate) and qty > 0 and rate > 0):
            raise ValueError(f"Item {index} quantity and rate must be greater than zero.")
        amount = qty * rate
        if not math.isfinite(amount):
            raise ValueError(f"Item {index} amount is too large.")
        items.append({
            "sl": index,
            "description": description,
            "hsn": hsn,
            "qty": qty,
            "rate": rate,
            "amount": round(amount, 2),
        })

    gst_type = payload.get("gst_type")
    if not isinstance(gst_type, str) or gst_type not in {"IGST", "CGST_SGST"}:
        raise ValueError("Choose IGST or CGST + SGST.")
    try:
        gst_rate = _parse_float(str(payload.get("gst_rate", "")))
    except (TypeError, ValueError) as exc:
        raise ValueError("GST rate must be a number.") from exc
    if not 0 <= gst_rate <= 100:
        raise ValueError("GST rate must be between 0 and 100.")

    taxable_value = round(sum(item["amount"] for item in items), 2)
    if not math.isfinite(taxable_value):
        raise ValueError("The combined item value is too large.")
    return {
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "delivery_note": _form_text(payload, "delivery_note", "Delivery note", limit=100) or "-",
        "payment_terms": _form_text(payload, "payment_terms", "Payment terms", limit=100) or "-",
        "buyer_order_no": _form_text(payload, "buyer_order_no", "Buyer order number", limit=100) or "-",
        "dispatch_doc_no": _form_text(payload, "dispatch_doc_no", "Dispatch document number", limit=100) or "-",
        "vehicle_no": _form_text(payload, "vehicle_no", "Vehicle number", limit=50) or "-",
        "destination": _form_text(payload, "destination", "Destination", limit=100) or "-",
        "delivery_terms": _form_text(payload, "delivery_terms", "Delivery terms", limit=100) or "As agreed",
        "buyer_name": _form_text(payload, "buyer_name", "Buyer name", required=True, limit=150),
        "buyer_address": _form_text(payload, "buyer_address", "Buyer address", required=True, limit=500),
        "buyer_gstin": _form_text(payload, "buyer_gstin", "Buyer GSTIN", limit=20),
        "items": items,
        "gst": calculate_gst(taxable_value, gst_rate, gst_type),
    }


async def web_app_invoice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    if not message or not message.web_app_data:
        return
    try:
        payload = json.loads(message.web_app_data.data)
        invoice_data = _invoice_from_web_app(payload, update)
    except (json.JSONDecodeError, ValueError) as exc:
        await message.reply_text(f"Form check: {exc} Please use /new to correct it.")
        return

    await _deliver_invoice(message, invoice_data)
    await message.reply_text(
        f"Excel aur PDF taiyaar hain. Invoice No. {invoice_data['invoice_no']}\n"
        "Naya invoice banane ke liye /new bhejein.",
        reply_markup=ReplyKeyboardRemove(),
    )


async def invoice_no(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if value.casefold() in {"auto", "-", "blank"}:
        context.user_data["invoice_no"] = None
    else:
        if not re.fullmatch(r"[A-Za-z0-9_-]{1,24}", value):
            await update.message.reply_text("Invoice number mein sirf letters, numbers, _ aur - use karein; ya AUTO bhejein.")
            return INVOICE_NO
        if not await _reserve_invoice_no(value):
            await update.message.reply_text("Is invoice number ki file pehle se hai. Dusra number ya AUTO bhejein.")
            return INVOICE_NO
        context.user_data["invoice_no"] = value
        await _record_manual_invoice_no(value)
    await update.message.reply_text("Invoice date? (DD-MM-YYYY, or send - for today)")
    return INVOICE_DATE


async def buyer_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["buyer_name"] = update.message.text.strip()
    await update.message.reply_text("Buyer's address?")
    return BUYER_ADDRESS


async def buyer_address(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["buyer_address"] = update.message.text.strip()
    await update.message.reply_text("Buyer's GSTIN? (send - to skip)")
    return BUYER_GSTIN


async def buyer_gstin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["buyer_gstin"] = None if text == "-" else text
    await update.message.reply_text("Buyer's Order No.? (send - to skip)")
    return BUYER_ORDER_NO


async def invoice_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["invoice_date"] = (
        date.today().strftime("%d-%m-%Y") if text == "-" else text
    )
    await update.message.reply_text("Buyer's name?")
    return BUYER_NAME


async def delivery_note(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["delivery_note"] = None if text == "-" else text
    await update.message.reply_text("Mode/Terms of Payment? (send - to skip)")
    return PAYMENT_TERMS


async def payment_terms(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["payment_terms"] = None if text == "-" else text
    await update.message.reply_text("Vehicle number? (send - to skip)")
    return VEHICLE_NO


async def buyer_order_no(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["buyer_order_no"] = None if text == "-" else text
    await update.message.reply_text("Dispatch Doc No.? (send - to skip)")
    return DISPATCH_DOC_NO


async def dispatch_doc_no(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["dispatch_doc_no"] = None if text == "-" else text
    await update.message.reply_text("Delivery Note? (send - to skip)")
    return DELIVERY_NOTE


async def vehicle_no(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["vehicle_no"] = None if text == "-" else text
    await update.message.reply_text("Destination city? (send - to skip)")
    return DESTINATION


async def destination(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["destination"] = None if text == "-" else text
    await update.message.reply_text("Terms of Delivery? (send - to use 'As agreed')")
    return DELIVERY_TERMS


async def delivery_terms(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    context.user_data["delivery_terms"] = "As agreed" if text == "-" else text
    await update.message.reply_text("Item description (e.g. METHI 460 Bags)?")
    return ITEM_DESC


async def item_desc(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["current_item"] = {"description": update.message.text.strip()}
    await update.message.reply_text("HSN/SAC code for this item?")
    return ITEM_HSN


async def item_hsn(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["current_item"]["hsn"] = update.message.text.strip()
    await update.message.reply_text("Quantity (e.g. 230)?")
    return ITEM_QTY


async def item_qty(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        qty = _parse_float(update.message.text)
    except ValueError:
        await update.message.reply_text("Please send a number for quantity, e.g. 230")
        return ITEM_QTY
    if qty <= 0:
        await update.message.reply_text("Quantity zero se zyada honi chahiye. Qtl mein quantity bhejein (12 ton = 120 Qtl).")
        return ITEM_QTY
    context.user_data["current_item"]["qty"] = qty
    await update.message.reply_text("Rate per Qtl (Rs.)?")
    return ITEM_RATE


async def item_rate(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        rate = _parse_float(update.message.text)
    except ValueError:
        await update.message.reply_text("Please send a number for the rate, e.g. 4000")
        return ITEM_RATE
    if rate <= 0:
        await update.message.reply_text("Rate zero se zyada honi chahiye. Per Qtl rate bhejein.")
        return ITEM_RATE

    item = context.user_data["current_item"]
    item["rate"] = rate
    item["amount"] = round(item["qty"] * rate, 2)
    item["sl"] = len(context.user_data["items"]) + 1
    context.user_data["items"].append(item)
    context.user_data.pop("current_item", None)

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("\u2795 Add another item", callback_data="add_item")],
        [InlineKeyboardButton("\u2705 Finish items", callback_data="finish_items")],
    ])
    await update.message.reply_text(
        f"Added: {item['description']} - Qty {item['qty']:g} x Rs. {item['rate']:,.2f} "
        f"= Rs. {item['amount']:,.2f}\n\nWhat next?",
        reply_markup=keyboard,
    )
    return ADD_MORE


async def add_more(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query.data == "add_item":
        if len(context.user_data["items"]) >= 11:
            await query.answer("This invoice format supports up to 11 items.", show_alert=True)
            return ADD_MORE
        await query.answer()
        await query.edit_message_text("Item description for the next item?")
        return ITEM_DESC

    await query.answer()
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("IGST (inter-state)", callback_data="IGST")],
        [InlineKeyboardButton("CGST + SGST (intra-state)", callback_data="CGST_SGST")],
    ])
    await query.edit_message_text("Which GST type applies?", reply_markup=keyboard)
    return GST_TYPE


async def gst_type(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    context.user_data["gst_type"] = query.data
    await query.edit_message_text("Total GST rate in %? (e.g. 5)")
    return GST_RATE


async def gst_rate(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        rate = _parse_float(update.message.text)
    except ValueError:
        await update.message.reply_text("Please send a number, e.g. 5")
        return GST_RATE
    if not 0 <= rate <= 100:
        await update.message.reply_text("GST rate 0 se 100 ke beech bhejein.")
        return GST_RATE

    ud = context.user_data
    taxable_value = round(sum(i["amount"] for i in ud["items"]), 2)
    gst = calculate_gst(taxable_value, rate, ud["gst_type"])

    invoice_data = {
        "invoice_no": ud.get("invoice_no") or await _next_invoice_no(),
        "invoice_date": ud["invoice_date"],
        "delivery_note": ud.get("delivery_note"),
        "payment_terms": ud.get("payment_terms"),
        "buyer_order_no": ud.get("buyer_order_no"),
        "dispatch_doc_no": ud.get("dispatch_doc_no"),
        "vehicle_no": ud.get("vehicle_no"),
        "destination": ud.get("destination"),
        "delivery_terms": ud.get("delivery_terms", "As agreed"),
        "buyer_name": ud["buyer_name"],
        "buyer_address": ud["buyer_address"],
        "buyer_gstin": ud.get("buyer_gstin"),
        "items": ud["items"],
        "gst": gst,
    }

    await _deliver_invoice(update.message, invoice_data)
    await update.message.reply_text("Excel aur PDF dono taiyaar hain. Naya invoice banane ke liye /start bhejein.")
    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Cancelled.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


async def report_handler_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Tell the user an update failed without exposing exception details/secrets."""
    error_name = type(context.error).__name__ if context.error else "UnknownError"
    logger.error("Telegram update handler failed (%s)", error_name)
    message = getattr(update, "effective_message", None)
    if message:
        try:
            await message.reply_text(
                "Is step par invoice process nahi ho paya. Apni details check karke "
                "dobara koshish karein, ya /new bhejkar form phir se shuru karein."
            )
        except Exception:
            logger.error("Could not send the invoice error notice to the user")


def build_application(token: str) -> Application:
    application = Application.builder().token(token).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start), CommandHandler("new", start)],
        states={
            INVOICE_NO: [MessageHandler(filters.TEXT & ~filters.COMMAND, invoice_no)],
            BUYER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, buyer_name)],
            BUYER_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, buyer_address)],
            BUYER_GSTIN: [MessageHandler(filters.TEXT & ~filters.COMMAND, buyer_gstin)],
            INVOICE_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, invoice_date)],
            DELIVERY_NOTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, delivery_note)],
            PAYMENT_TERMS: [MessageHandler(filters.TEXT & ~filters.COMMAND, payment_terms)],
            BUYER_ORDER_NO: [MessageHandler(filters.TEXT & ~filters.COMMAND, buyer_order_no)],
            DISPATCH_DOC_NO: [MessageHandler(filters.TEXT & ~filters.COMMAND, dispatch_doc_no)],
            VEHICLE_NO: [MessageHandler(filters.TEXT & ~filters.COMMAND, vehicle_no)],
            DESTINATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, destination)],
            DELIVERY_TERMS: [MessageHandler(filters.TEXT & ~filters.COMMAND, delivery_terms)],
            ITEM_DESC: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_desc)],
            ITEM_HSN: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_hsn)],
            ITEM_QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_qty)],
            ITEM_RATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, item_rate)],
            ADD_MORE: [CallbackQueryHandler(add_more)],
            GST_TYPE: [CallbackQueryHandler(gst_type)],
            GST_RATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, gst_rate)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    application.add_handler(conv_handler)
    application.add_handler(
        MessageHandler(filters.StatusUpdate.WEB_APP_DATA, web_app_invoice)
    )
    application.add_error_handler(report_handler_error)
    return application


def main() -> None:
    # Loads TELEGRAM_BOT_TOKEN from a .env file placed next to this script
    # (see .env.example). An OS environment variable, if set, takes priority.
    load_dotenv()
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise SystemExit(
            "No bot token found. Create a file named '.env' next to main.py "
            "containing:\n    TELEGRAM_BOT_TOKEN=your-token-here\n"
            "(see .env.example for a template)."
        )

    # Python 3.14 removed the implicit event loop that older versions of
    # python-telegram-bot (21.x) rely on in run_polling(). Create one
    # explicitly so the bot starts correctly on Python 3.12-3.14+.
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())

    application = build_application(token)
    logger.info("Bot starting...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

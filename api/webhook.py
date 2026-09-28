"""Vercel Python Function endpoint for Telegram webhook updates."""

import asyncio
import json
import logging
import os
from http.server import BaseHTTPRequestHandler

from telegram import Update

from main import build_application

logger = logging.getLogger(__name__)
logger.setLevel(logging.WARNING)


async def _process_update(payload: dict) -> None:
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    application = build_application(token)
    initialized = False
    try:
        await application.initialize()
        initialized = True
        update = Update.de_json(payload, application.bot)
        await application.process_update(update)
    finally:
        if initialized:
            await application.shutdown()


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        expected_secret = os.environ.get("TELEGRAM_WEBHOOK_SECRET", "")
        provided_secret = self.headers.get("X-Telegram-Bot-Api-Secret-Token", "")
        if not expected_secret or provided_secret != expected_secret:
            self.send_error(403, "Invalid webhook secret")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            if content_length <= 0 or content_length > 1_000_000:
                self.send_error(413, "Invalid update size")
                return
            payload = json.loads(self.rfile.read(content_length))
            asyncio.run(_process_update(payload))
        except Exception as exc:
            # Exception strings from Telegram/httpx can contain request URLs,
            # and Telegram Bot API URLs embed the bot token. Never log them.
            logger.error("Webhook update failed (%s)", type(exc).__name__)
            self.send_error(500, "Update processing failed")
            return

        self.send_response(200)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"Invoice bot webhook is ready")

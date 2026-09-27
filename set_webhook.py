"""Point Telegram to the deployed Vercel webhook endpoint."""

import os
import sys

import httpx
from dotenv import load_dotenv


def main() -> None:
    load_dotenv()
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    secret = os.environ.get("TELEGRAM_WEBHOOK_SECRET")
    if not token or not secret:
        raise SystemExit("Set TELEGRAM_BOT_TOKEN and TELEGRAM_WEBHOOK_SECRET in .env first.")
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python set_webhook.py https://your-project.vercel.app")

    base_url = sys.argv[1].rstrip("/")
    response = httpx.post(
        f"https://api.telegram.org/bot{token}/setWebhook",
        json={
            "url": f"{base_url}/api/webhook",
            "secret_token": secret,
            "allowed_updates": ["message", "callback_query"],
        },
        timeout=20,
    )
    response.raise_for_status()
    result = response.json()
    if not result.get("ok"):
        raise SystemExit(f"Telegram rejected the webhook: {result.get('description', 'unknown error')}")
    print(f"Webhook registered: {base_url}/api/webhook")


if __name__ == "__main__":
    main()

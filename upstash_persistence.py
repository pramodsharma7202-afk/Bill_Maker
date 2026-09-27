"""PTB persistence backed by Upstash Redis REST for serverless webhooks."""

import json
import os
from collections import defaultdict
from urllib.parse import quote

import httpx
from telegram.ext import BasePersistence


class UpstashPersistence(BasePersistence):
    """Persist Telegram conversation/user data across Vercel invocations."""

    def __init__(self) -> None:
        super().__init__(update_interval=0)
        self.url = os.environ["UPSTASH_REDIS_REST_URL"].rstrip("/")
        self.token = os.environ["UPSTASH_REDIS_REST_TOKEN"]
        self.prefix = "bill-maker:ptb:"

    async def _command(self, *parts: object, value: str | None = None):
        path = "/".join(quote(str(part), safe="") for part in parts)
        headers = {"Authorization": f"Bearer {self.token}"}
        async with httpx.AsyncClient(timeout=8) as client:
            if value is None:
                response = await client.get(f"{self.url}/{path}", headers=headers)
            else:
                response = await client.post(
                    f"{self.url}/{path}", headers=headers, content=value
                )
        response.raise_for_status()
        result = response.json()
        if "error" in result:
            raise RuntimeError(f"Upstash Redis error: {result['error']}")
        return result.get("result")

    async def _hash(self, key: str) -> dict[str, str]:
        result = await self._command("hgetall", key)
        if not result:
            return {}
        if isinstance(result, dict):
            return {str(k): str(v) for k, v in result.items()}
        return {str(result[i]): str(result[i + 1]) for i in range(0, len(result), 2)}

    async def _get_dict(self, key: str):
        values = await self._hash(key)
        return defaultdict(dict, {int(k): json.loads(v) for k, v in values.items()})

    async def _set_hash_item(self, key: str, field: object, value: object) -> None:
        await self._command(
            "hset", key, field, value=json.dumps(value, ensure_ascii=False)
        )

    async def _drop_hash_item(self, key: str, field: object) -> None:
        await self._command("hdel", key, field)

    async def get_user_data(self):
        return await self._get_dict(self.prefix + "user-data")

    async def get_chat_data(self):
        return await self._get_dict(self.prefix + "chat-data")

    async def get_bot_data(self):
        value = await self._command("get", self.prefix + "bot-data")
        return json.loads(value) if value else {}

    async def get_callback_data(self):
        return None

    async def get_conversations(self, name: str):
        values = await self._hash(self.prefix + "conversation:" + name)
        conversations = {}
        for key, state in values.items():
            decoded_key = tuple(json.loads(key))
            conversations[decoded_key] = json.loads(state)
        return conversations

    async def update_user_data(self, user_id: int, data) -> None:
        await self._set_hash_item(self.prefix + "user-data", user_id, data)

    async def update_chat_data(self, chat_id: int, data) -> None:
        await self._set_hash_item(self.prefix + "chat-data", chat_id, data)

    async def update_bot_data(self, data) -> None:
        await self._command(
            "set", self.prefix + "bot-data", value=json.dumps(data, ensure_ascii=False)
        )

    async def update_callback_data(self, data) -> None:
        return None

    async def update_conversation(self, name: str, key: tuple, new_state) -> None:
        redis_key = self.prefix + "conversation:" + name
        field = json.dumps(key, separators=(",", ":"))
        if new_state is None:
            await self._drop_hash_item(redis_key, field)
        else:
            await self._set_hash_item(redis_key, field, new_state)

    async def refresh_user_data(self, user_id: int, user_data) -> None:
        return None

    async def refresh_chat_data(self, chat_id: int, chat_data) -> None:
        return None

    async def refresh_bot_data(self, bot_data) -> None:
        return None

    async def drop_user_data(self, user_id: int) -> None:
        await self._drop_hash_item(self.prefix + "user-data", user_id)

    async def drop_chat_data(self, chat_id: int) -> None:
        await self._drop_hash_item(self.prefix + "chat-data", chat_id)

    async def flush(self) -> None:
        return None

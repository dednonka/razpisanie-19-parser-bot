"""
Версия бота, которая ходит в Telegram через MTProto-прокси (например, локальный
tg-ws-proxy-android на 127.0.0.1:1443). Обычный Bot API через MTProto-прокси не работает,
поэтому здесь используется библиотека Telethon, которая говорит на MTProto.

Файл лежит рядом с schedule_bot.py и использует его парсер и экраны.

Установка:  pip install telethon requests openpyxl
Переменные окружения (задаватйте командой export):
  Запуск:$env:BOT_TOKEN="токен"
  $env:API_ID="..."
  $env:API_HASH="..."
  $env:MTPROXY_SECRET="ключ прокси"
  $env:MTPROXY_PORT="порт"
  python schedule_bot_mtproto.py
"""
import asyncio
import logging
import os
import sys

from telethon import Button, TelegramClient, connection, events
from telethon.errors import MessageNotModifiedError

import schedule_bot as sb

log = logging.getLogger("schedule.mtproto")


def need(name):
    v = os.getenv(name)
    if not v:
        sys.exit(f"Не задана переменная {name}")
    return v


API_ID = int(need("API_ID"))
API_HASH = need("API_HASH")
SECRET = need("MTPROXY_SECRET")
HOST = os.getenv("MTPROXY_HOST", "127.0.0.1")
PORT = int(os.getenv("MTPROXY_PORT", "1443"))
BOT_TOKEN = need("BOT_TOKEN")

client = TelegramClient(
    "bot_session",
    API_ID,
    API_HASH,
    connection=connection.ConnectionTcpMTProxyRandomizedIntermediate,
    proxy=(HOST, PORT, SECRET),
)


def to_buttons(markup):
    if not markup or not markup.get("inline_keyboard"):
        return None
    return [
        [Button.inline(b["text"], data=b["callback_data"].encode()) for b in row]
        for row in markup["inline_keyboard"]
    ]


@client.on(events.NewMessage(pattern=r"^/(start|update)"))
async def on_message(event):
    if event.raw_text.startswith("/start"):
        text, markup = sb.screen_dates()
        if markup:
            text = "Привет! 👋 Я показываю расписание уроков.\n\n" + text
        await event.respond(text, buttons=to_buttons(markup), parse_mode="html")
    else:
        await asyncio.to_thread(sb.sync_refresh)
        await event.respond(
            f"Обновлено. Дат в базе: {len(sb.SCHEDULES)}, файлов-мастеров: {len(sb.WEEK_FILES)}"
        )


@client.on(events.CallbackQuery)
async def on_callback(event):
    kind, *a = event.data.decode().split(":")
    if kind == "back":
        res = sb.screen_dates()
    elif kind == "find":
        res = sb.screen_find_grades()
    elif kind == "hol":
        res = sb.screen_holidays()
    elif kind == "fg":
        res = sb.screen_find_letters(a[0])
    elif kind == "fl":
        res = sb.screen_find_subjects(a[0], a[1])
    elif kind == "fs":
        res = sb.screen_next_lesson(a[0], a[1], a[2])
    elif kind == "wk":
        res = sb.screen_weekdays(a[0])
    elif kind == "d":
        res = sb.screen_grades(a[0])
    elif kind == "g":
        res = sb.screen_letters(a[0], a[1])
    elif kind == "c":
        res = sb.screen_schedule(a[0], a[1], a[2])
    else:
        res = None
    if res is None:
        await event.answer("Это расписание уже удалили, нажми /start", alert=True)
        return
    text, markup = res
    try:
        await event.edit(text, buttons=to_buttons(markup), parse_mode="html")
    except MessageNotModifiedError:
        pass
    await event.answer()


async def refresher():
    while True:
        try:
            await asyncio.to_thread(sb.sync_refresh)
        except Exception:
            log.exception("Не удалось обновить папку")
        await asyncio.sleep(sb.REFRESH_SECONDS)


async def main():
    await client.start(bot_token=BOT_TOKEN)
    me = await client.get_me()
    log.info("Бот запущен через MTProto-прокси: @%s", me.username)
    asyncio.create_task(refresher())
    await client.run_until_disconnected()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    asyncio.run(main())

"""
Версия бота, которая ходит в Telegram через MTProto-прокси (например, локальный
tg-ws-proxy-android на 127.0.0.1:1443). Обычный Bot API через MTProto-прокси не работает,
поэтому здесь используется библиотека Telethon, которая говорит на MTProto.

Файл лежит рядом с schedule_bot.py и использует его парсер и экраны.

Установка:  pip install telethon requests openpyxl
Переменные окружения (задавай в Termux командой export):
  BOT_TOKEN        токен бота от @BotFather
  API_ID           число с my.telegram.org  (раздел API development tools)
  API_HASH         строка оттуда же
  MTPROXY_SECRET   секрет прокси из приложения (hex, из ссылки tg://proxy ... secret=...)
  MTPROXY_HOST     по умолчанию 127.0.0.1
  MTPROXY_PORT     по умолчанию 1443
  MTPROXY_LINK     вместо трёх переменных выше можно дать целую ссылку
                   tg://proxy?server=...&port=...&secret=...
Автосмена прокси: когда прокси перестаёт отвечать, бот сам берёт свежие ссылки
tg://proxy?... из канала-источника, проверяет их по очереди, подставляет рабочие
хост/порт/секрет и запоминает их в proxy_state.json (после перезапуска начнёт с них).
  PROXY_SOURCE         канал со ссылками на прокси (по умолчанию https://t.me/MTProxy4free)
  PROXY_AUTOSWITCH     0 — выключить автосмену (по умолчанию включена)
  PROXY_CHECK_INTERVAL как часто проверять живость прокси, секунд (по умолчанию 30)
  PROXY_STATE_FILE     файл с запомненным прокси (по умолчанию proxy_state.json)
  PARSING_MTPRTO   канал, который надо парсить через тот же прокси, например @dednonka
                   (или ссылка t.me/dednonka). Не задан — парсер канала выключен.
  PARSING_LIMIT    сколько последних постов держать в памяти (по умолчанию 20)
  PARSING_INTERVAL как часто проверять канал, секунд (по умолчанию 300)
  PARSING_SESSION  имя файла сессии парсера (по умолчанию parser_session)
Парсер канала читает историю канала, а бот (Bot API) этого делать не может, поэтому
он заходит под ТВОИМ аккаунтом: при самом первом запуске Telethon спросит номер
телефона и код из Telegram, дальше вход сохраняется в parser_session.session.
Команда в боте: /channel (или /news) — показать последние посты канала.
Запуск:  python schedule_bot_mtproto.py
"""
import asyncio
import base64
import html as htmllib
import json
import logging
import os
import re
import sys
from urllib.parse import parse_qs, unquote, urlparse

import requests

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
BOT_TOKEN = need("BOT_TOKEN")


def _channel_name(raw):
    """@dednonka / dednonka / https://t.me/dednonka -> 'dednonka' (None, если не задан)."""
    raw = (raw or "").strip()
    if not raw:
        return None
    raw = re.sub(r"^(https?://)?(t\.me|telegram\.me)/(s/)?", "", raw)
    return raw.split("/")[0].split("?")[0].lstrip("@") or None


# ───────────────────── прокси: разбор ссылок и автосмена ─────────────────────

PROXY_SOURCE = os.getenv("PROXY_SOURCE", "https://t.me/MTProxy4free")
PROXY_AUTOSWITCH = os.getenv("PROXY_AUTOSWITCH", "1") != "0"
PROXY_CHECK_INTERVAL = int(os.getenv("PROXY_CHECK_INTERVAL", "30"))
PROXY_STATE_FILE = os.getenv("PROXY_STATE_FILE", "proxy_state.json")

_PROXY_LINK_RE = re.compile(r"(?:tg://|https?://t\.me/)proxy\?[^\s\"'<>\)\]]+", re.I)


def _norm_secret(raw):
    """Секрет в ссылке бывает hex или base64url — приводим к hex."""
    raw = unquote(raw or "").strip()
    if re.fullmatch(r"[0-9a-fA-F]+", raw) and len(raw) % 2 == 0:
        return raw.lower()
    try:
        return base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4)).hex() or None
    except Exception:
        return None


def parse_proxy_link(link):
    """tg://proxy?server=H&port=P&secret=S  (или https://t.me/proxy?...) -> (H, P, S) либо None."""
    link = htmllib.unescape((link or "").strip())
    q = parse_qs(urlparse(link).query)
    try:
        host = q["server"][0].strip()
        port = int(q["port"][0])
        secret = _norm_secret(q["secret"][0])
    except (KeyError, IndexError, ValueError):
        return None
    if not host or not secret or not 0 < port < 65536:
        return None
    return (host, port, secret)


def find_proxies(text):
    """Достаёт ВСЕ ссылки на прокси из произвольного текста/HTML, без повторов, в порядке появления."""
    out = []
    for m in _PROXY_LINK_RE.finditer(htmllib.unescape(text or "")):
        p = parse_proxy_link(m.group(0))
        if p and p not in out:
            out.append(p)
    return out


def _fetch_proxies_web():
    """Читает публичную веб-версию канала (t.me/s/<канал>) обычным HTTPS. Новые прокси — первыми."""
    name = _channel_name(PROXY_SOURCE)
    if not name:
        return []
    kw = {}
    if os.getenv("TG_PROXY"):
        kw["proxies"] = {"http": os.environ["TG_PROXY"], "https": os.environ["TG_PROXY"]}
    r = requests.get(f"https://t.me/s/{name}", headers=sb.HEADERS, timeout=15, **kw)
    r.raise_for_status()
    return list(reversed(find_proxies(r.text)))     # на странице новые посты внизу


def _load_state():
    try:
        with open(PROXY_STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _save_state():
    try:
        with open(PROXY_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"current": list(CURRENT), "known": [list(p) for p in KNOWN[:30]]},
                      f, ensure_ascii=False, indent=1)
    except OSError:
        log.exception("Не удалось записать %s", PROXY_STATE_FILE)


def _env_proxy():
    link = os.getenv("MTPROXY_LINK")
    if link:
        p = parse_proxy_link(link)
        if p:
            return p
        sys.exit("MTPROXY_LINK задан, но это не ссылка tg://proxy?server=...&port=...&secret=...")
    secret = os.getenv("MTPROXY_SECRET")
    if secret:
        return (os.getenv("MTPROXY_HOST", "127.0.0.1"), int(os.getenv("MTPROXY_PORT", "1443")), secret)
    return None


def _tup(p):
    return (str(p[0]), int(p[1]), str(p[2]).lower())


ENV_PROXY = _env_proxy()
_state = _load_state()
KNOWN = [_tup(p) for p in _state.get("known", []) if len(p) == 3]
if _state.get("current"):                 # последний рабочий — с него и начинаем
    CURRENT = _tup(_state["current"])
elif ENV_PROXY:
    CURRENT = _tup(ENV_PROXY)
else:                                     # ничего не задано — берём первый из канала-источника
    try:
        _first = _fetch_proxies_web()
    except Exception:
        _first = []
    if not _first:
        sys.exit("Не задан прокси (MTPROXY_LINK или MTPROXY_SECRET), а из канала PROXY_SOURCE "
                 "достать его не удалось.")
    CURRENT = _tup(_first[0])
HOST, PORT, SECRET = CURRENT

client = TelegramClient(
    "bot_session",
    API_ID,
    API_HASH,
    connection=connection.ConnectionTcpMTProxyRandomizedIntermediate,
    proxy=CURRENT,
    timeout=10,
    connection_retries=2,
    retry_delay=1,
)

BAD: dict = {}                     # прокси -> когда не заработал (не пробуем повторно 10 минут)
ROTATE_LOCK = None                 # создаётся в main(), внутри event loop
parser_ready = False               # запущен ли клиент парсера канала


async def tcp_alive(host, port, timeout=5):
    """Быстрая предпроверка: принимает ли хост:порт TCP-соединение вообще."""
    try:
        _, w = await asyncio.wait_for(asyncio.open_connection(host, port), timeout)
        w.close()
        return True
    except Exception:
        return False


async def _fetch_proxies_via_telethon():
    """Запасной путь, если t.me по HTTPS недоступен: читаем канал через клиента парсера."""
    if not (parser_client and parser_ready):
        return []
    found = []
    async for m in parser_client.iter_messages(_channel_name(PROXY_SOURCE), limit=30):
        chunks = [m.message or ""]
        chunks += [getattr(e, "url", "") or "" for e in (m.entities or [])]
        for row in (m.buttons or []):
            chunks += [getattr(b, "url", "") or "" for b in row]
        for p in find_proxies("\n".join(chunks)):
            if p not in found:
                found.append(p)
    return found


async def collect_candidates():
    """Кого пробовать: исходный прокси из env, свежие из канала, затем ранее запомненные."""
    fresh = []
    try:
        fresh = await asyncio.to_thread(_fetch_proxies_web)
    except Exception as e:
        log.warning("Не удалось открыть канал с прокси по HTTPS (%s)", e.__class__.__name__)
    if not fresh:
        try:
            fresh = await _fetch_proxies_via_telethon()
        except Exception:
            log.exception("Не удалось прочитать канал с прокси через Telethon")
    for p in fresh:
        if p not in KNOWN:
            KNOWN.insert(0, p)
    cands = ([_tup(ENV_PROXY)] if ENV_PROXY else []) + [_tup(p) for p in fresh] + list(KNOWN)
    seen, out = set(), []
    for p in cands:
        if p not in seen and p != CURRENT:
            seen.add(p)
            out.append(p)
    # FakeTLS-секреты (начинаются с ee) Telethon поддерживает плохо — пробуем их в последнюю очередь
    out.sort(key=lambda p: p[2].startswith("ee"))
    return out


async def apply_proxy(p):
    """Переключает клиентов на прокси p и проверяет, что связь реально есть."""
    from telethon.tl import functions
    for c in (client, parser_client if parser_ready else None):
        if c is not None:
            c.set_proxy(p)
    for c in (client, parser_client if parser_ready else None):
        if c is not None:
            try:
                await asyncio.wait_for(c.disconnect(), 10)
            except Exception:
                pass
    await asyncio.wait_for(client.connect(), 30)
    await asyncio.wait_for(client(functions.help.GetConfigRequest()), 20)
    if parser_ready:
        try:
            await asyncio.wait_for(parser_client.connect(), 30)
        except Exception:
            log.exception("Клиент парсера не переподключился после смены прокси")


async def rotate_proxy(reason=""):
    """Подбирает рабочий прокси из канала-источника. True, если нашёлся и применён."""
    global CURRENT, HOST, PORT, SECRET
    async with ROTATE_LOCK:
        log.warning("Меняю прокси%s", f" ({reason})" if reason else "")
        now = asyncio.get_running_loop().time()
        tried = 0
        for p in await collect_candidates():
            if now - BAD.get(p, -1e9) < 600:
                continue
            if not await tcp_alive(p[0], p[1]):
                BAD[p] = now
                continue
            tried += 1
            if tried > 15:
                break
            log.info("Пробую прокси %s:%s", p[0], p[1])
            try:
                await apply_proxy(p)
            except Exception as e:
                log.info("Не подошёл %s:%s (%s)", p[0], p[1], e.__class__.__name__)
                BAD[p] = now
                continue
            CURRENT = p
            HOST, PORT, SECRET = p
            if p not in KNOWN:
                KNOWN.insert(0, p)
            _save_state()
            log.warning("Теперь работаю через %s:%s", p[0], p[1])
            return True
        log.error("Рабочий прокси не нашёлся, попробую позже")
        return False


async def proxy_watchdog():
    """Раз в PROXY_CHECK_INTERVAL секунд проверяет связь; после двух провалов подряд меняет прокси."""
    from telethon.tl import functions
    fails = 0
    while True:
        await asyncio.sleep(PROXY_CHECK_INTERVAL)
        try:
            await asyncio.wait_for(client(functions.help.GetConfigRequest()), 20)
            fails = 0
        except Exception as e:
            fails += 1
            log.warning("Нет ответа через %s:%s (%s), подряд: %d", CURRENT[0], CURRENT[1],
                        e.__class__.__name__, fails)
            if fails >= 2 and await rotate_proxy("прокси не отвечает"):
                fails = 0


async def connect_with_failover():
    """Первый вход бота. Если стартовый прокси мёртв — сразу ищем живой."""
    while True:
        try:
            await asyncio.wait_for(client.start(bot_token=BOT_TOKEN), 60)
            return
        except Exception as e:
            if not PROXY_AUTOSWITCH:
                raise
            log.warning("Не удалось подключиться через %s:%s (%s)", CURRENT[0], CURRENT[1],
                        e.__class__.__name__)
            try:
                await asyncio.wait_for(client.disconnect(), 10)
            except Exception:
                pass
            if not await rotate_proxy("при запуске"):
                await asyncio.sleep(30)


# ───────────────────── парсер Telegram-канала через тот же прокси ─────────────────────

PARSING_CHANNEL = _channel_name(os.getenv("PARSING_MTPRTO"))
PARSING_LIMIT = int(os.getenv("PARSING_LIMIT", "20"))
PARSING_INTERVAL = int(os.getenv("PARSING_INTERVAL", "300"))
PARSING_SESSION = os.getenv("PARSING_SESSION", "parser_session")

CHANNEL_POSTS: list = []     # последние посты, новые первыми: {id, date, text, media, link}

parser_client = None
if PARSING_CHANNEL:
    # отдельная пользовательская сессия, но тот же MTProto-прокси, что и у бота
    parser_client = TelegramClient(
        PARSING_SESSION,
        API_ID,
        API_HASH,
        connection=connection.ConnectionTcpMTProxyRandomizedIntermediate,
        proxy=CURRENT,
    )


async def parse_channel(limit=None):
    """Читает последние посты канала PARSING_MTPRTO и обновляет CHANNEL_POSTS.

    Возвращает список НОВЫХ постов (которых не было при прошлой проверке).
    """
    if not parser_client:
        return []
    limit = limit or PARSING_LIMIT
    posts = []
    async for m in parser_client.iter_messages(PARSING_CHANNEL, limit=limit):
        text = (m.message or "").strip()
        media = type(m.media).__name__.replace("MessageMedia", "") if m.media else ""
        if not text and not media:
            continue                                   # служебные сообщения
        posts.append({
            "id": m.id,
            "date": m.date.astimezone(sb.EKB_TZ) if m.date else None,
            "text": text,
            "media": media,                            # Photo / Document / ... или ""
            "link": f"https://t.me/{PARSING_CHANNEL}/{m.id}",
        })
    known = {p["id"] for p in CHANNEL_POSTS}
    new = [p for p in posts if p["id"] not in known] if CHANNEL_POSTS else []
    CHANNEL_POSTS[:] = posts
    if new:
        log.info("Канал @%s: новых постов %d", PARSING_CHANNEL, len(new))
    return new


def format_posts(posts, count=5, max_len=600):
    if not PARSING_CHANNEL:
        return "Парсер канала выключен: задай PARSING_MTPRTO=@канал и перезапусти бота."
    if not posts:
        return f"В @{PARSING_CHANNEL} пока нет постов (или канал ещё не прочитан)."
    out = [f"<b>@{htmllib.escape(PARSING_CHANNEL)}</b> — последние посты\n"]
    for p in posts[:count]:
        when = p["date"].strftime("%d.%m %H:%M") if p["date"] else ""
        body = p["text"]
        if len(body) > max_len:
            body = body[:max_len].rstrip() + "…"
        if p["media"] and not body:
            body = f"[{p['media']}]"
        elif p["media"]:
            body += f"\n[{p['media']}]"
        out.append(f'🗓 {when}  <a href="{p["link"]}">пост</a>\n{htmllib.escape(body)}\n')
    return "\n".join(out)


def to_buttons(markup):
    if not markup or not markup.get("inline_keyboard"):
        return None
    return [
        [Button.inline(b["text"], data=b["callback_data"].encode()) for b in row]
        for row in markup["inline_keyboard"]
    ]


@client.on(events.NewMessage(pattern=r"^/(start|update|channel|news)"))
async def on_message(event):
    try:
        if event.raw_text.startswith("/start"):
            # screen_dates() ходит за погодой в сеть — уводим с event loop,
            # чтобы бот не вис на всех остальных во время запроса/таймаута
            text, markup = await asyncio.to_thread(sb.screen_dates)
            if markup:
                text = "Привет! 👋 Я показываю расписание уроков.\n\n" + text
            await event.respond(text, buttons=to_buttons(markup), parse_mode="html")
        elif re.match(r"^/(channel|news)", event.raw_text):
            if parser_client:
                try:
                    await parse_channel()
                except Exception:
                    log.exception("Не удалось прочитать канал, показываю сохранённое")
            await event.respond(format_posts(CHANNEL_POSTS), parse_mode="html", link_preview=False)
        else:
            await asyncio.to_thread(sb.sync_refresh)
            await event.respond(
                f"Обновлено. Дат в базе: {len(sb.SCHEDULES)}, файлов-мастеров: {len(sb.WEEK_FILES)}"
            )
    except Exception:
        log.exception("Ошибка обработки сообщения")
        try:
            await event.respond("Что-то пошло не так, попробуй ещё раз чуть позже.")
        except Exception:
            pass


@client.on(events.CallbackQuery)
async def on_callback(event):
    try:
        kind, *a = event.data.decode().split(":")
        if kind == "back":
            res = await asyncio.to_thread(sb.screen_dates)
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
    except Exception:
        log.exception("Ошибка обработки нажатия кнопки")
        try:
            await event.answer("Что-то пошло не так, попробуй /start", alert=True)
        except Exception:
            pass


async def refresher():
    while True:
        try:
            await asyncio.to_thread(sb.sync_refresh)
        except Exception:
            log.exception("Не удалось обновить папку")
        await asyncio.sleep(sb.REFRESH_SECONDS)


async def channel_refresher():
    while True:
        try:
            await parse_channel()
        except Exception:
            log.exception("Не удалось прочитать канал @%s", PARSING_CHANNEL)
        await asyncio.sleep(PARSING_INTERVAL)


async def main():
    global ROTATE_LOCK, parser_ready
    ROTATE_LOCK = asyncio.Lock()
    await connect_with_failover()
    me = await client.get_me()
    log.info("Бот запущен через MTProto-прокси %s:%s: @%s", CURRENT[0], CURRENT[1], me.username)
    asyncio.create_task(refresher())
    if parser_client:
        # при первом запуске спросит телефон и код; потом вход берётся из файла сессии
        await parser_client.start()
        parser_ready = True
        log.info("Парсер канала @%s запущен", PARSING_CHANNEL)
        asyncio.create_task(channel_refresher())
    if PROXY_AUTOSWITCH:
        asyncio.create_task(proxy_watchdog())
    # run_until_disconnected() здесь нельзя: при смене прокси клиент переподключается,
    # и оно бы завершило программу. Просто живём вечно.
    await asyncio.Event().wait()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    asyncio.run(main())

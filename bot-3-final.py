import asyncio
import logging
import os
import re
import sqlite3
from datetime import datetime, timedelta
from io import BytesIO

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, BufferedInputFile, CallbackQuery
from aiogram.utils.keyboard import InlineKeyboardBuilder
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# НАСТРОЙКИ
# ---------------------------------------------------------------------------

BOT_TOKEN = os.getenv("BOT_TOKEN", "ВАШ_ТОКЕН_СЮДА")
DB_PATH = os.getenv("DB_PATH", "swim.db")

# Telegram numeric ID (не username!) владельца и тренера — только у них есть
# право добавлять/менять результаты и смотреть общую таблицу.
# Узнать свой ID можно написав боту @userinfobot.
ADMIN_IDS = {
    5220385313,  # <- впиши сюда свой Telegram ID
    5683235845,}

logging.basicConfig(level=logging.INFO)

STROKE_ALIASES = {
    "кроль": "кроль", "вольный": "кроль", "вс": "кроль", "вольный стиль": "кроль", "кр": "кроль",
    "брасс": "брасс","бр": "брасс",
    "спина": "спина", "на спине": "спина", "сп": "спина",
    "батт": "баттерфляй", "баттерфляй": "баттерфляй", "дельфин": "баттерфляй", "бт": "баттерфляй",
    "компл": "комплекс", "комплекс": "комплекс", "к/п": "комплекс", "кп": "комплекс", 
}
# Ожидающие подтверждения имени записи: telegram_id тренера -> (record, warnings)
PENDING_RECORDS = {}
# Ожидающий ввод тренировок по дням: trainer_id -> {dates: [...], index: int}
PENDING_TRAININGS = {}


def is_admin(telegram_id: int) -> bool:
    return telegram_id in ADMIN_IDS

# ---------------------------------------------------------------------------
# БАЗА ДАННЫХ
# ---------------------------------------------------------------------------

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            distance INTEGER NOT NULL,
            stroke TEXT NOT NULL,
            time_seconds REAL NOT NULL,
            session_type TEXT NOT NULL DEFAULT 'тренировка',
            pool INTEGER NOT NULL DEFAULT 25,
            swim_date TEXT,
            training_kind TEXT,
            body_part TEXT,
            fins INTEGER NOT NULL DEFAULT 0,
            paddles TEXT NOT NULL DEFAULT 'нет',
            snorkel INTEGER NOT NULL DEFAULT 0,
            recorded_at TEXT NOT NULL
        )
        """
    )
    # мягкая миграция для баз, созданных более ранней версией бота
    existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(results)")}
    migrations = {
        "pool": "ALTER TABLE results ADD COLUMN pool INTEGER NOT NULL DEFAULT 25",
        "swim_date": "ALTER TABLE results ADD COLUMN swim_date TEXT",
        "training_kind": "ALTER TABLE results ADD COLUMN training_kind TEXT",
        "body_part": "ALTER TABLE results ADD COLUMN body_part TEXT",
        "fins": "ALTER TABLE results ADD COLUMN fins INTEGER NOT NULL DEFAULT 0",
        "paddles": "ALTER TABLE results ADD COLUMN paddles TEXT NOT NULL DEFAULT 'нет'",
        "snorkel": "ALTER TABLE results ADD COLUMN snorkel INTEGER NOT NULL DEFAULT 0",
    }
    for col, ddl in migrations.items():
        if col not in existing_cols:
            conn.execute(ddl)
    conn.execute("UPDATE results SET swim_date = substr(recorded_at, 1, 10) WHERE swim_date IS NULL")

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            telegram_id INTEGER PRIMARY KEY,
            name TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS standards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            distance INTEGER NOT NULL,
            stroke TEXT NOT NULL,
            pool INTEGER NOT NULL DEFAULT 25,
            label TEXT NOT NULL,
            time_seconds REAL NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS trainings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            training_date TEXT NOT NULL UNIQUE,
            title TEXT NOT NULL DEFAULT 'Тренировка',
            description TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def link_user(telegram_id, name):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO users (telegram_id, name) VALUES (?, ?) "
        "ON CONFLICT(telegram_id) DO UPDATE SET name=excluded.name",
        (telegram_id, name),
    )
    conn.commit()
    conn.close()


def get_linked_name(telegram_id):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT name FROM users WHERE telegram_id = ?", (telegram_id,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row else None


def get_all_linked_users():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT telegram_id, name FROM users ORDER BY name")
    rows = cur.fetchall()
    conn.close()
    return rows


def unlink_user(telegram_id):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("DELETE FROM users WHERE telegram_id = ?", (telegram_id,))
    conn.commit()
    changed = cur.rowcount > 0
    conn.close()
    return changed


def get_known_names():
    conn = sqlite3.connect(DB_PATH)
    names = set()
    for (n,) in conn.execute("SELECT DISTINCT name FROM results"):
        names.add(n)
    for (n,) in conn.execute("SELECT DISTINCT name FROM users"):
        names.add(n)
    conn.close()
    return sorted(names)


def save_result(record):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO results (name, distance, stroke, time_seconds, session_type, pool, swim_date, "
        "training_kind, body_part, fins, paddles, snorkel, recorded_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            record["name"], record["distance"], record["stroke"], record["time_seconds"],
            record["session_type"], record["pool"], record["swim_date"],
            record["training_kind"], record["body_part"],
            record["fins"], record["paddles"], record["snorkel"],
            datetime.now().isoformat(),
        ),
    )
    conn.commit()
    conn.close()


def update_result(result_id, record):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "UPDATE results SET name=?, distance=?, stroke=?, time_seconds=?, session_type=?, pool=?, "
        "swim_date=?, training_kind=?, body_part=?, fins=?, paddles=?, snorkel=? WHERE id=?",
        (
            record["name"], record["distance"], record["stroke"], record["time_seconds"],
            record["session_type"], record["pool"], record["swim_date"],
            record["training_kind"], record["body_part"],
            record["fins"], record["paddles"], record["snorkel"], result_id,
        ),
    )
    conn.commit()
    changed = cur.rowcount > 0
    conn.close()
    return changed


def delete_result(result_id):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("DELETE FROM results WHERE id = ?", (result_id,))
    conn.commit()
    changed = cur.rowcount > 0
    conn.close()
    return changed


def reindex_results():
    """Пересчитывает id оставшихся записей подряд: 1, 2, 3... без пропусков."""
    conn = sqlite3.connect(DB_PATH)
    cols = (
        "name, distance, stroke, time_seconds, session_type, pool, swim_date, "
        "training_kind, body_part, fins, paddles, snorkel, recorded_at"
    )
    rows = conn.execute(f"SELECT {cols} FROM results ORDER BY id").fetchall()
    conn.execute("DELETE FROM results")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='results'")
    placeholders = ",".join(["?"] * 13)
    conn.executemany(f"INSERT INTO results ({cols}) VALUES ({placeholders})", rows)
    conn.commit()
    conn.close()


def find_matching_results(name, distance, stroke, time_seconds, session_type=None):
    conn = sqlite3.connect(DB_PATH)
    query = (
        "SELECT id, session_type, swim_date, pool FROM results "
        "WHERE name = ? AND distance = ? AND stroke = ? AND ABS(time_seconds - ?) < 0.01"
    )
    params = [name, distance, stroke, time_seconds]
    if session_type:
        query += " AND session_type = ?"
        params.append(session_type)
    cur = conn.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def get_previous_result(name, distance, stroke):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT time_seconds, swim_date, session_type FROM results "
        "WHERE name = ? AND distance = ? AND stroke = ? "
        "ORDER BY swim_date DESC, recorded_at DESC LIMIT 2",
        (name, distance, stroke),
    )
    rows = cur.fetchall()
    conn.close()
    if len(rows) >= 2:
        return rows[1]
    return None


def get_personal_best(name, distance, stroke):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT MIN(time_seconds) FROM results WHERE name = ? AND distance = ? AND stroke = ?",
        (name, distance, stroke),
    )
    row = cur.fetchone()
    conn.close()
    return row[0] if row else None


def get_distances_for_name(name):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT DISTINCT distance, stroke FROM results WHERE name = ? ORDER BY distance, stroke",
        (name,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_entries(name, distance, stroke, session_type, limit=2):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT time_seconds, swim_date FROM results "
        "WHERE name = ? AND distance = ? AND stroke = ? AND session_type = ? "
        "ORDER BY swim_date DESC, recorded_at DESC LIMIT ?",
        (name, distance, stroke, session_type, limit),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_best(name, distance, stroke, session_type):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT MIN(time_seconds) FROM results "
        "WHERE name = ? AND distance = ? AND stroke = ? AND session_type = ?",
        (name, distance, stroke, session_type),
    )
    row = cur.fetchone()
    conn.close()
    return row[0] if row else None


def get_best_with_pool(name, distance, stroke, session_type):
    """Как get_best, но заодно возвращает бассейн этого лучшего результата —
    нужно, чтобы сравнивать со «своими» нормативами (они тоже привязаны к бассейну)."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT time_seconds, pool FROM results "
        "WHERE name = ? AND distance = ? AND stroke = ? AND session_type = ? "
        "ORDER BY time_seconds ASC LIMIT 1",
        (name, distance, stroke, session_type),
    )
    row = cur.fetchone()
    conn.close()
    return row if row else None


def get_all_results(name=None):
    conn = sqlite3.connect(DB_PATH)
    query = (
        "SELECT name, distance, stroke, time_seconds, session_type, pool, swim_date, "
        "training_kind, body_part, fins, paddles, snorkel, recorded_at FROM results "
    )
    params = ()
    if name:
        query += "WHERE name = ? "
        params = (name,)
    query += "ORDER BY name, distance, stroke, swim_date, recorded_at"
    cur = conn.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def get_recent(limit=15):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT id, name, distance, stroke, time_seconds, session_type, pool, swim_date "
        "FROM results ORDER BY id DESC LIMIT ?",
        (limit,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows

# ---------------------------------------------------------------------------
# ПЛАН ТРЕНИРОВОК
# ---------------------------------------------------------------------------

def parse_date_value(raw):
    raw = raw.strip()
    m = re.match(r"^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?$", raw)
    if not m:
        return None
    day, month, year = m.groups()
    year = year or str(datetime.now().year)
    if len(year) == 2:
        year = "20" + year
    try:
        return datetime(int(year), int(month), int(day)).strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_training_dates(raw):
    raw = raw.strip().replace("—", "-").replace("–", "-")
    if "," in raw:
        dates = [parse_date_value(x) for x in raw.split(",") if x.strip()]
        return dates if dates and all(dates) else None
    if "-" in raw:
        left, right = [x.strip() for x in raw.split("-", 1)]
        start, end = parse_date_value(left), parse_date_value(right)
        if not start or not end:
            return None
        a, b = datetime.strptime(start, "%Y-%m-%d"), datetime.strptime(end, "%Y-%m-%d")
        if b < a or (b-a).days > 366:
            return None
        return [(a + timedelta(days=i)).strftime("%Y-%m-%d") for i in range((b-a).days+1)]
    one = parse_date_value(raw)
    return [one] if one else None


def save_training(training_date, title, description):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO trainings (training_date, title, description, created_at) VALUES (?, ?, ?, ?) "
        "ON CONFLICT(training_date) DO UPDATE SET title=excluded.title, description=excluded.description",
        (training_date, title.strip() or "Тренировка", description.strip(), datetime.now().isoformat()),
    )
    conn.commit(); conn.close()


def get_training(training_date):
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute("SELECT id, training_date, title, description FROM trainings WHERE training_date = ?", (training_date,)).fetchone()
    conn.close(); return row


def delete_training(training_date):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("DELETE FROM trainings WHERE training_date = ?", (training_date,))
    conn.commit(); changed = cur.rowcount > 0; conn.close(); return changed


def get_training_dates():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT training_date, title FROM trainings ORDER BY training_date").fetchall()
    conn.close(); return rows


def format_training(training):
    if not training:
        return None
    _, training_date, title, description = training
    date_label = datetime.strptime(training_date, "%Y-%m-%d").strftime("%d.%m.%Y")
    return f"🏊 <b>{title}</b>\n📅 {date_label}\n\n{description}"


# ---------------------------------------------------------------------------
# НОРМАТИВЫ
# ---------------------------------------------------------------------------

def add_standard(distance, stroke, pool, label, time_seconds):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO standards (distance, stroke, pool, label, time_seconds) VALUES (?, ?, ?, ?, ?)",
        (distance, stroke, pool, label, time_seconds),
    )
    conn.commit()
    conn.close()


def delete_standard(std_id):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("DELETE FROM standards WHERE id = ?", (std_id,))
    conn.commit()
    changed = cur.rowcount > 0
    conn.close()
    return changed


def get_standards(distance=None, stroke=None, pool=None):
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT id, distance, stroke, pool, label, time_seconds FROM standards WHERE 1=1"
    params = []
    if distance is not None:
        query += " AND distance = ?"
        params.append(distance)
    if stroke is not None:
        query += " AND stroke = ?"
        params.append(stroke)
    if pool is not None:
        query += " AND pool = ?"
        params.append(pool)
    query += " ORDER BY distance, stroke, pool, time_seconds DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return rows


def get_standard_distances(pool=None):
    conn = sqlite3.connect(DB_PATH)
    if pool is None:
        rows = conn.execute("SELECT DISTINCT distance FROM standards ORDER BY distance").fetchall()
    else:
        rows = conn.execute(
            "SELECT DISTINCT distance FROM standards WHERE pool = ? ORDER BY distance", (pool,)
        ).fetchall()
    conn.close()
    return [r[0] for r in rows]


def get_standard_strokes(pool=None, distance=None):
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT DISTINCT stroke FROM standards WHERE 1=1"
    params = []
    if pool is not None:
        query += " AND pool = ?"
        params.append(pool)
    if distance is not None:
        query += " AND distance = ?"
        params.append(distance)
    query += " ORDER BY stroke"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [r[0] for r in rows]


def format_standards_message(rows):
    lines = ["<b>Нормативы</b>:"]
    last_key = None
    for std_id, d, s, pool, label, time_seconds in rows:
        key = (d, s, pool)
        if key != last_key:
            lines.append(f"\n<b>{d}м {s}, бассейн {pool}м:</b>")
            last_key = key
        lines.append(f"  #{std_id} {label} — {format_time(time_seconds)}")
    return "\n".join(lines)


def get_norm_status(distance, stroke, pool, time_seconds):
    """Сравнивает результат с нормативами на эту дистанцию/стиль/бассейн.
    Возвращает None, если нормативов для такой дистанции/стиля/бассейна нет.
    Иначе — словарь: achieved_all=True и top_label (лучший выполненный), либо
    achieved_all=False, label/target_time/gap — ближайший невыполненный."""
    rows = get_standards(distance, stroke, pool)
    if not rows:
        return None
    not_met = [r for r in rows if time_seconds > r[5]]
    if not not_met:
        best = min(rows, key=lambda r: r[5])
        return {"achieved_all": True, "top_label": best[4]}
    target = max(not_met, key=lambda r: r[5])
    return {
        "achieved_all": False,
        "label": target[4],
        "target_time": target[5],
        "gap": time_seconds - target[5],
    }

# ---------------------------------------------------------------------------
# ПАРСИНГ
# ---------------------------------------------------------------------------

def parse_time(raw):
    raw = raw.strip().replace(",", ".")
    if "." in raw:
        main, frac = raw.rsplit(".", 1)
        frac_seconds = float("0." + frac)
    else:
        main = raw
        frac_seconds = 0.0
    parts = [p for p in main.split(":") if p != ""]
    if len(parts) == 1:
        total = float(parts[0])
    elif len(parts) == 2:
        m, s = parts
        total = int(m) * 60 + float(s)
    elif len(parts) == 3:
        a, b, c = parts
        if int(c) < 100 and "." not in raw:
            total = int(a) * 60 + int(b) + int(c) / 100
        else:
            total = int(a) * 3600 + int(b) * 60 + float(c)
    else:
        raise ValueError(f"не смог разобрать время: {raw}")
    return total + frac_seconds


def format_time(seconds):
    m = int(seconds // 60)
    s = seconds - m * 60
    if m > 0:
        return f"{m}:{s:05.2f}"
    return f"{s:.2f}"


def format_delta(delta):
    if delta is None:
        return "—"
    if abs(delta) < 0.005:
        return "0.00"
    sign = "-" if delta < 0 else "+"
    return f"{sign}{format_time(abs(delta))}"


def normalize_stroke(raw):
    key = raw.strip().lower()
    return STROKE_ALIASES.get(key, raw.strip())


def parse_stroke_field(raw):
    """Разбирает поле «Стиль».

    Понимает обычные названия стилей, а также сокращение для заплывов
    ногами:
      - "ноги"        -> стиль по умолчанию "кроль", часть тела "ноги"
      - "ноги брасс"  -> стиль "брасс", часть тела "ноги"
    Возвращает (stroke, body_part_override), где body_part_override — None,
    если поле было обычным названием стиля.
    """
    s = raw.strip().lower()
    if s in ("ноги", "ног"):
        return "кроль", "ноги"
    for prefix in ("ноги ", "ног "):
        if s.startswith(prefix):
            rest = s[len(prefix):].strip()
            if rest:
                return normalize_stroke(rest), "ноги"
    return normalize_stroke(raw), None


# Двусловные названия стилей (напр. "на спине", "вольный стиль") — нужны,
# чтобы при разборе записи без тире не разорвать их пробелом на два поля.
MULTIWORD_STROKES = {k for k in STROKE_ALIASES if " " in k}


def merge_multiword_tokens(words):
    """Склеивает соседние слова, которые при записи без тире должны остаться
    одним полем: "лопатки мал", "ноги брасс", "на спине", "без инвентаря" и т.п.
    """
    segments = []
    i, n = 0, len(words)
    while i < n:
        w = words[i]
        wl = w.lower()
        nxt = words[i + 1] if i + 1 < n else ""
        nxtl = nxt.lower()

        if wl == "бассейн" and re.match(r"^\d+м?$", nxtl):
            segments.append(f"{w} {nxt}")
            i += 2
            continue
        if "лопат" in wl and (nxtl.startswith("мал") or nxtl.startswith("бол")):
            segments.append(f"{w} {nxt}")
            i += 2
            continue
        if wl in ("ноги", "ног") and nxtl in STROKE_ALIASES:
            segments.append(f"{w} {nxt}")
            i += 2
            continue
        if nxt and f"{wl} {nxtl}" in MULTIWORD_STROKES:
            segments.append(f"{w} {nxt}")
            i += 2
            continue
        if wl == "без" and nxtl == "инвентаря":
            segments.append(f"{w} {nxt}")
            i += 2
            continue

        segments.append(w)
        i += 1
    return segments


def split_fields(text):
    """Разбивает текст записи на поля.

    Поддерживает и классический формат через тире:
        Имя - Дистанция - Стиль - Время - [доп. части в любом порядке]
    и запись просто через пробелы, вообще без тире:
        Имя Дистанция Стиль Время [доп. части в любом порядке]
    Первые четыре поля (имя, дистанция, стиль, время) в обоих случаях идут
    строго в этом порядке; всё, что после них — доп. части — можно писать в
    любой последовательности, как и раньше.
    """
    cleaned = text.replace("—", "-").replace("–", "-")
    if "-" in cleaned:
        return [p.strip() for p in cleaned.split("-") if p.strip()]
    return merge_multiword_tokens(cleaned.split())


def parse_equipment(segment, stroke):
    s = segment.lower()
    fins = 1 if "ласт" in s else 0
    paddles = "нет"
    if "лопат" in s:
        if "мал" in s:
            paddles = "мал"
        elif "бол" in s:
            paddles = "больш"
        else:
            paddles = "мал" if stroke == "брасс" else "больш"
    snorkel = 1 if "труб" in s else 0
    return fins, paddles, snorkel


def equipment_str(fins, paddles, snorkel):
    parts = []
    if fins:
        parts.append("ласты")
    if paddles and paddles != "нет":
        parts.append(f"лопатки ({paddles})")
    if snorkel:
        parts.append("труба")
    return "+".join(parts) if parts else "без"


def classify_segment(seg):
    """Определяет, к какому полю относится доп. часть записи (дата/бассейн/тип и т.д.)."""
    s = seg.strip()
    sl = s.lower()

    m = re.match(r"^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?$", s)
    if m:
        day, month, year = m.groups()
        year = year or str(datetime.now().year)
        if len(year) == 2:
            year = "20" + year
        try:
            date_obj = datetime(int(year), int(month), int(day))
            return "date", date_obj.strftime("%Y-%m-%d")
        except ValueError:
            pass

    if "соревы" in sl:
        return "session_type", "соревнования"
    if "треня" in sl:
        return "session_type", "тренировка"
    m = re.match(r"^бассейн\s*(\d+)\s*м?$", sl)
    if m:
        return "pool", int(m.group(1))
    if "тест" in sl:
        return "training_kind", "тест"
    if "серия" in sl:
        return "training_kind", "серия"
    if "к/р" in sl:
        return "body_part", "координация"
    if "рук" in sl:
        return "body_part", "руки"
    if "ног" in sl:
        return "body_part", "ноги"
    if any(k in sl for k in ("ласт", "лопат", "труб")):
        return "equipment", sl
    if sl in ("без", "нет", "без инвентаря", "ничего"):
        return "equipment", sl
    return None, seg


def parse_record_full(text):
    """Полный парсер для добавления записи. Формат:
    Имя - Дистанция - Стиль - Время [- доп.части в любом порядке]
    Возвращает (record_dict, warnings) или (None, []) если не распознал базовые поля.
    """
    parts = split_fields(text)
    if len(parts) < 4:
        return None, []
    name = parts[0].title()
    distance_digits = re.sub(r"[^\d]", "", parts[1])
    if not distance_digits:
        return None, []
    distance = int(distance_digits)
    stroke, body_part_override = parse_stroke_field(parts[2])
    try:
        time_seconds = parse_time(parts[3])
    except ValueError:
        return None, []

    record = {
        "name": name, "distance": distance, "stroke": stroke, "time_seconds": time_seconds,
        "session_type": "тренировка", "pool": 25,
        "swim_date": datetime.now().strftime("%Y-%m-%d"),
        "training_kind": "тест", "body_part": "координация",
        "fins": 0, "paddles": "нет", "snorkel": 0,
    }
    if body_part_override:
        record["body_part"] = body_part_override
    warnings = []
    for seg in parts[4:]:
        kind, value = classify_segment(seg)
        if kind == "date":
            record["swim_date"] = value
        elif kind == "session_type":
            record["session_type"] = value
        elif kind == "pool":
            record["pool"] = value
        elif kind == "training_kind":
            record["training_kind"] = value
        elif kind == "body_part":
            record["body_part"] = value
        elif kind == "equipment":
            fins, paddles, snorkel = parse_equipment(value, stroke)
            record["fins"], record["paddles"], record["snorkel"] = fins, paddles, snorkel
        else:
            warnings.append(seg)
    return record, warnings


def parse_basic(text):
    """Короткий парсер для команды /delete по содержимому: Имя - Дистанция - Стиль - Время [- тип]."""
    parts = split_fields(text)
    if len(parts) < 4:
        return None
    name = parts[0].title()
    distance_digits = re.sub(r"[^\d]", "", parts[1])
    if not distance_digits:
        return None
    distance = int(distance_digits)
    stroke, _ = parse_stroke_field(parts[2])
    try:
        time_seconds = parse_time(parts[3])
    except ValueError:
        return None
    session_type = None
    if len(parts) >= 5:
        t = parts[4].lower()
        if "сорев" in t:
            session_type = "соревнования"
        elif "трен" in t:
            session_type = "тренировка"
    return name, distance, stroke, time_seconds, session_type


def parse_batch_header(header_text):
    """Разбирает общую часть пакетной записи — всё, что относится сразу к
    нескольким пловцам: дистанция, стиль/часть тела, инвентарь, режим, тип,
    дата, бассейн. Возвращает (record_template без имени и времени, warnings).
    """
    cleaned = header_text.replace("—", "-").replace("–", "-")
    if "-" in cleaned:
        segments = [p.strip() for p in cleaned.split("-") if p.strip()]
    else:
        segments = merge_multiword_tokens(cleaned.split())

    record = {
        "distance": None, "stroke": "кроль",
        "session_type": "тренировка", "pool": 25,
        "swim_date": datetime.now().strftime("%Y-%m-%d"),
        "training_kind": "тест", "body_part": "координация",
        "fins": 0, "paddles": "нет", "snorkel": 0,
    }
    warnings = []
    stroke_set = False
    for seg in segments:
        sl = seg.strip().lower()

        m = re.match(r"^(\d+)\s*м?$", sl)
        if m and record["distance"] is None:
            record["distance"] = int(m.group(1))
            continue

        if not stroke_set:
            stroke, body_part_override = parse_stroke_field(seg)
            if body_part_override or sl in STROKE_ALIASES:
                record["stroke"] = stroke
                if body_part_override:
                    record["body_part"] = body_part_override
                stroke_set = True
                continue

        kind, value = classify_segment(seg)
        if kind == "date":
            record["swim_date"] = value
        elif kind == "session_type":
            record["session_type"] = value
        elif kind == "pool":
            record["pool"] = value
        elif kind == "training_kind":
            record["training_kind"] = value
        elif kind == "body_part":
            record["body_part"] = value
        elif kind == "equipment":
            fins, paddles, snorkel = parse_equipment(value, record["stroke"])
            record["fins"], record["paddles"], record["snorkel"] = fins, paddles, snorkel
        else:
            warnings.append(seg)
    return record, warnings


def _make_batch_records(template, header_warnings, items):
    records = []
    for item in items:
        item_clean = re.sub(r"\s*-\s*", " ", item).strip()
        pieces = item_clean.rsplit(None, 1)
        if len(pieces) != 2:
            return None
        name_part, time_part = pieces
        try:
            time_seconds = parse_time(time_part)
        except ValueError:
            return None
        record = dict(template)
        record["name"] = name_part.strip().title()
        record["time_seconds"] = time_seconds
        records.append((record, list(header_warnings)))
    return records or None


def parse_batch_records(text):
    """Пакетная запись с ':' или без него.

    Например: «50м ноги труба тест : Аня 40; Оля 30» и
    «50м ноги труба тест Аня 40».
    """
    cleaned = text.strip()
    if not cleaned:
        return None

    if ":" in cleaned:
        header_text, _, body_text = cleaned.partition(":")
        header_text, body_text = header_text.strip(), body_text.strip()
        if not header_text or not body_text:
            return None
        items = [p.strip() for p in re.split(r"[;\n]+", body_text) if p.strip()]
        template, header_warnings = parse_batch_header(header_text)
        if template["distance"] is None:
            return None
        return _make_batch_records(template, header_warnings, items)

    words = cleaned.split()
    if len(words) < 5:
        return None

    candidates = []
    for split_at in range(3, len(words)-1):
        header_text = " ".join(words[:split_at])
        body_text = " ".join(words[split_at:])
        template, header_warnings = parse_batch_header(header_text)
        if template["distance"] is None:
            continue
        pieces = body_text.rsplit(None, 1)
        if len(pieces) != 2:
            continue
        name_part, time_part = pieces
        try:
            parse_time(time_part)
        except ValueError:
            continue
        if name_part.strip():
            candidates.append((split_at, template, header_warnings, [body_text]))

    if not candidates:
        return None
    _, template, header_warnings, items = max(candidates, key=lambda x: x[0])
    return _make_batch_records(template, header_warnings, items)

# ---------------------------------------------------------------------------
# БОТ
# ---------------------------------------------------------------------------

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

ADMIN_HELP = (
    "🏊 <b>Режим тренера/владельца</b>\n\n"
    "<b>Добавить результат</b> — сообщение в формате:\n"
    "<code>Имя - Дистанция - Стиль - Время [- доп. части]</code>\n"
    "Тире можно вообще не писать, просто через пробел: "
    "<code>Аня 200 брасс 3:58.00 15.03 50 сорев</code> — Имя, Дистанция, Стиль и Время "
    "должны идти именно в этом порядке, а доп. части (дата, бассейн, тип и т.д.) — "
    "в любом порядке, с тире или без.\n"
    "Для заплыва одними ногами не нужно повторять стиль дважды: вместо «кроль ... ноги» "
    "можно просто написать <code>ноги</code> в поле стиля — это будет кроль ногами; "
    "а для другого стиля — <code>ноги брасс</code>, <code>ноги спина</code> и т.п.\n"
    "<b>Пакетом на нескольких пловцов</b> — общая часть (дистанция, стиль/ноги, "
    "инвентарь, режим и т.д.), потом «:» и список «Имя время» через «;»:\n"
    "<code>50м ноги труба тест : Аня 40; Оля 30</code>\n"
    "Проверка «это новый пловец?» в пакетном режиме не показывается — имена "
    "сохраняются как есть, поэтому проверяй написание.\n"
    "Доп. части можно писать в любом порядке, бот сам поймёт что где:\n"
    "• дата заплыва: <code>15.03</code> или <code>15.03.2026</code> (по умолчанию — сегодня)\n"
    "• бассейн: обычно можно не писать вообще (по умолчанию 25м); если нужен "
    "50-метровый — напиши <code>бассейн 50</code>\n"
    "• тип: <code>треня</code> / <code>соревы</code> (по умолчанию треня)\n"
    "• для трень — характер: <code>координация</code>/<code>руки</code>/<code>ноги</code> "
    "(по умолчанию координация)\n"
    "• для трень — режим: <code>серия</code>/<code>тест</code> (по умолчанию тест)\n"
    "• для трень — инвентарь: <code>ласты</code>/<code>лопатки</code>/<code>труба</code> и их "
    "комбинации через «+» (по умолчанию без). Лопатки можно уточнить: "
    "<code>лопатки мал</code> / <code>лопатки бол</code> (без уточнения — брасс маленькие, "
    "остальные стили большие)\n\n"
    "Пример: <code>Аня - 200 - брасс - 3:58.00 - 15.03 - 50 - сорев</code>\n\n"
    "Если имя не совпадёт ни с одним уже известным пловцом — бот не сохранит, а предложит "
    "выбрать из существующих или подтвердить, что это новый пловец.\n\n"
    "Если для дистанции/стиля есть нормативы (см. /normativy), бот сразу покажет, сколько "
    "не хватает до ближайшего невыполненного, либо что все нормативы уже выполнены.\n\n"
    "<b>Команды</b> (у каждой есть русский алиас с «!», можно с аргументами так же, как у /команды):\n"
    "/add ... или !добавить ... — то же самое явной командой (для групп, см. /help ниже про приватность)\n"
    "/last или !последние — последние 15 записей с их id\n"
    "/delete 15 или !удалить 15 — удалить запись №15\n"
    "/delete 15 16 20 — удалить сразу несколько записей по id\n"
    "/delete Аня - 200 - брасс - 4:00.00 - соревы — удалить по содержимому (если совпадение "
    "не одно — покажет варианты с id)\n"
    "/edit 15 - Аня - 200 - брасс - 3:58.00 - соревы или !редакт 15 - ... — исправить запись №15 (формат как у /add)\n"
    "/table или !таблица — Excel со всеми результатами всех пловцов\n"
    "/table Имя — Excel по одному пловцу\n"
    "/addnorm 100 - брасс - 25 - 3 юн - 1:45.00 или !добавнорм ... — добавить норматив (бассейн можно "
    "не указывать, по умолчанию 25м; можно сразу несколько строк одним сообщением)\n"
    "/normativy или !нормативы — все нормативы; !нормативы 100 брасс — только по дистанции/стилю\n"
    "/delnorm 5 или !удалнорм 5 — удалить норматив по id\n"
    "/link id Имя или !привязать id Имя — привязать чужой Telegram ID к имени пловца\n"
    "/unlink id или !отвязать id — отвязать\n"
    "/users или !пользователи — список привязанных пловцов\n"
    "/my или !результат — посмотреть свой личный прогресс (как у обычных пользователей)\n"
    "/addtraining 16.09-20.09 или !добавтреню 16.09-20.09 — создать разные тренировки по дням\nПосле команды бот попросит тренировку отдельно для каждого дня\n"
    "/training или !треня — тренировка на сегодня (в сообщении есть кнопка «Нормативы»)\n"
    "/prevtraining или !предтреня — предыдущая доступная тренировка\n"
    "/trainings или !тренировки — список тренировок (тренер)\n"
    "/deltraining 18.09 или !удалтреню 18.09 — удалить тренировку на дату (тренер)\n"
    "/whoami или !ктоя — твой Telegram ID и роль\n"
)

USER_HELP = (
    "🏊 Привет! Здесь можно смотреть только свои результаты.\n\n"
    "Если тренер ещё не привязал тебя, привяжи себя сам:\n"
    "<code>/iam Твоё Имя</code> или <code>!яэто Твоё Имя</code>\n\n"
    "После этого команда /my (или !результат) покажет кнопки с твоими дистанциями — жми на "
    "нужную, и бот пришлёт лучший результат на соревнованиях, лучший на "
    "тренировках и два последних результата в каждой из категорий.\n\n"
    "<code>/normativy</code> или <code>!нормативы</code> — посмотреть все нормативы, "
    "<code>!нормативы 100 брасс</code> — нормативы на конкретную дистанцию и стиль.\n\n"
    "<code>/training</code> или <code>!треня</code> — тренировка на сегодня. Если её нет — «выходной».\n"
    "<code>/prevtraining</code> или <code>!предтреня</code> — предыдущая тренировка."
)


@dp.message(CommandStart())
async def cmd_start(message: Message):
    if is_admin(message.from_user.id):
        await message.answer(ADMIN_HELP, parse_mode="HTML")
    else:
        await message.answer(USER_HELP, parse_mode="HTML")


@dp.message(Command("help"))
async def cmd_help(message: Message):
    await cmd_start(message)


@dp.message(Command("whoami"))
async def cmd_whoami(message: Message):
    uid = message.from_user.id
    name = get_linked_name(uid)
    role = "тренер/владелец" if is_admin(uid) else "пловец"
    text = f"Твой Telegram ID: <code>{uid}</code>\nРоль: {role}\n"
    text += f"Привязанное имя: {name}" if name else "Имя не привязано (используй /iam Имя)"
    await message.answer(text, parse_mode="HTML")


@dp.message(Command("iam"))
async def cmd_iam(message: Message):
    arg = message.text.partition(" ")[2].strip()
    if not arg:
        await message.answer("Напиши так: /iam Аня")
        return
    name = arg.title()
    link_user(message.from_user.id, name)
    await message.answer(f"Готово! Теперь ты — {name}. Жми /my, чтобы увидеть свой прогресс.")


@dp.message(Command("link"))
async def cmd_link(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    arg = message.text.partition(" ")[2].strip()
    telegram_id_str, _, name = arg.partition(" ")
    if not telegram_id_str.isdigit() or not name.strip():
        await message.answer(
            "Напиши так: /link 123456789 Аня\n"
            "(узнать чужой Telegram ID: пусть человек напишет боту @userinfobot "
            "или пришлёт тебе свой /whoami)"
        )
        return
    name = name.strip().title()
    link_user(int(telegram_id_str), name)
    await message.answer(f"Готово! ID {telegram_id_str} привязан к имени «{name}».")


@dp.message(Command("unlink"))
async def cmd_unlink(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    arg = message.text.partition(" ")[2].strip()
    if not arg.isdigit():
        await message.answer("Напиши так: /unlink 123456789")
        return
    ok = unlink_user(int(arg))
    await message.answer("Отвязано ✅" if ok else "Такой ID не привязан ни к кому.")


@dp.message(Command("users"))
async def cmd_users(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    rows = get_all_linked_users()
    if not rows:
        await message.answer("Пока никто не привязан. Используй /link id Имя.")
        return
    lines = ["<b>Привязанные пользователи</b>:"]
    for telegram_id, name in rows:
        lines.append(f"• {name} — <code>{telegram_id}</code>")
    await message.answer("\n".join(lines), parse_mode="HTML")

# ---------------------------------------------------------------------------
# ДОБАВЛЕНИЕ РЕЗУЛЬТАТОВ (только тренер/владелец)
# ---------------------------------------------------------------------------

def build_confirmation(record, warnings, prev, best_before):
    name, distance, stroke = record["name"], record["distance"], record["stroke"]
    time_seconds = record["time_seconds"]

    lines = [
        f"✅ Записано: <b>{name}</b> — {distance}м {stroke}, "
        f"{format_time(time_seconds)}, бассейн {record['pool']}м, {record['swim_date']} "
        f"({record['session_type']})"
    ]
    if record["session_type"] == "тренировка":
        lines.append(
            f"{record['training_kind']}, {record['body_part']}, "
            f"инвентарь: {equipment_str(record['fins'], record['paddles'], record['snorkel'])}"
        )

    if prev:
        delta = time_seconds - prev[0]
        if delta < 0:
            lines.append(f"⬇️ Быстрее предыдущего на {format_time(abs(delta))}")
        elif delta > 0:
            lines.append(f"⬆️ Медленнее предыдущего на {format_time(abs(delta))}")
        else:
            lines.append("➖ Точно как в прошлый раз")
    else:
        lines.append("📌 Первый результат на этой дистанции этим стилем")

    if best_before is not None:
        if time_seconds < best_before:
            lines.append("🏆 Новый личный рекорд!")
        else:
            gap = time_seconds - best_before
            lines.append(f"Личный рекорд: {format_time(best_before)} (отставание {format_time(gap)})")

    if warnings:
        lines.append("⚠️ Не понял и проигнорировал: " + ", ".join(f"«{w}»" for w in warnings))

    status = get_norm_status(distance, stroke, record["pool"], time_seconds)
    if status:
        if status["achieved_all"]:
            lines.append(f"🥇 Выполнены все нормативы на эту дистанцию (лучший: {status['top_label']})")
        else:
            lines.append(
                f"📐 До норматива «{status['label']}» ({format_time(status['target_time'])}) "
                f"не хватает {format_time(status['gap'])}"
            )

    return "\n".join(lines)


async def finalize_save(record, warnings, message):
    prev = get_previous_result(record["name"], record["distance"], record["stroke"])
    best_before = get_personal_best(record["name"], record["distance"], record["stroke"])
    save_result(record)
    text = build_confirmation(record, warnings, prev, best_before)
    await message.answer(text, parse_mode="HTML")


async def handle_batch_text(message: Message, batch):
    known_names = get_known_names()
    header = batch[0][0]
    header_warnings = batch[0][1]

    lines = ["✅ Записано пакетом:"]
    lines.append(
        f"{header['distance']}м {header['stroke']}, {header['body_part']}, "
        f"{header['training_kind']}, инвентарь: "
        f"{equipment_str(header['fins'], header['paddles'], header['snorkel'])}, "
        f"бассейн {header['pool']}м, {header['swim_date']} ({header['session_type']})"
    )

    new_names = []
    for record, _ in batch:
        if known_names and record["name"] not in known_names and record["name"] not in new_names:
            new_names.append(record["name"])

        prev = get_previous_result(record["name"], record["distance"], record["stroke"])
        best_before = get_personal_best(record["name"], record["distance"], record["stroke"])
        save_result(record)

        delta_str = ""
        if prev:
            delta = record["time_seconds"] - prev[0]
            if delta < -0.005:
                delta_str = f" (⬇️ {format_time(abs(delta))})"
            elif delta > 0.005:
                delta_str = f" (⬆️ {format_time(abs(delta))})"
            else:
                delta_str = " (➖)"
        pr_str = " 🏆" if best_before is not None and record["time_seconds"] < best_before else ""
        norm_str = ""
        status = get_norm_status(record["distance"], record["stroke"], record["pool"], record["time_seconds"])
        if status:
            if status["achieved_all"]:
                norm_str = f" [🥇 все нормативы, лучший: {status['top_label']}]"
            else:
                norm_str = f" [до «{status['label']}» не хватает {format_time(status['gap'])}]"
        lines.append(f"• {record['name']} — {format_time(record['time_seconds'])}{delta_str}{pr_str}{norm_str}")

    if new_names:
        lines.append("⚠️ Новые имена (проверь написание): " + ", ".join(new_names))
    if header_warnings:
        lines.append("⚠️ Не понял в общей части: " + ", ".join(f"«{w}»" for w in header_warnings))

    await message.answer("\n".join(lines), parse_mode="HTML")


async def handle_incoming_text(message: Message, raw_text: str):
    batch = parse_batch_records(raw_text)
    if batch:
        await handle_batch_text(message, batch)
        return

    record, warnings = parse_record_full(raw_text)
    if not record:
        await message.answer(
            "Не смог разобрать запись 🤔\nФормат: <code>Имя - Дистанция - Стиль - Время</code>",
            parse_mode="HTML",
        )
        return

    known_names = get_known_names()
    if known_names and record["name"] not in known_names:
        PENDING_RECORDS[message.from_user.id] = (record, warnings)
        builder = InlineKeyboardBuilder()
        for n in known_names:
            builder.button(text=n, callback_data=f"usename:{n}")
        builder.button(text=f"➕ Новый пловец «{record['name']}»", callback_data="newname")
        builder.button(text="❌ Отмена", callback_data="cancelrecord")
        builder.adjust(2)
        await message.answer(
            f"Не нашёл пловца «{record['name']}» в базе. Выбери из существующих или "
            f"подтверди, что это новый пловец:",
            reply_markup=builder.as_markup(),
        )
        return

    await finalize_save(record, warnings, message)


@dp.callback_query(F.data.startswith("usename:"))
async def on_use_existing_name(callback: CallbackQuery):
    admin_id = callback.from_user.id
    pending = PENDING_RECORDS.pop(admin_id, None)
    if not pending:
        await callback.answer("Эта запись уже неактуальна, отправь заново.", show_alert=True)
        return
    record, warnings = pending
    _, chosen_name = callback.data.split(":", 1)
    record["name"] = chosen_name
    await finalize_save(record, warnings, callback.message)
    await callback.answer()


@dp.callback_query(F.data == "newname")
async def on_confirm_new_name(callback: CallbackQuery):
    admin_id = callback.from_user.id
    pending = PENDING_RECORDS.pop(admin_id, None)
    if not pending:
        await callback.answer("Эта запись уже неактуальна, отправь заново.", show_alert=True)
        return
    record, warnings = pending
    await finalize_save(record, warnings, callback.message)
    await callback.answer()


@dp.callback_query(F.data == "cancelrecord")
async def on_cancel_record(callback: CallbackQuery):
    PENDING_RECORDS.pop(callback.from_user.id, None)
    await callback.message.answer("Отменено, ничего не записано.")
    await callback.answer()


@dp.message(Command("add"))
async def cmd_add(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Добавлять результаты может только тренер или владелец.")
        return
    text_after_command = message.text.partition(" ")[2]
    if not text_after_command.strip():
        await message.answer("Напиши так: /add Аня - 200 - брасс - 4:00.00")
        return
    await handle_incoming_text(message, text_after_command)

# ---------------------------------------------------------------------------
# ПРОСМОТР / ИСПРАВЛЕНИЕ / УДАЛЕНИЕ (только тренер/владелец)
# ---------------------------------------------------------------------------

@dp.message(Command("last"))
async def cmd_last(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    rows = get_recent(15)
    if not rows:
        await message.answer("Записей пока нет.")
        return
    lines = ["<b>Последние записи</b>:"]
    for rid, name, distance, stroke, time_seconds, session_type, pool, swim_date in rows:
        lines.append(
            f"#{rid} — {name}, {distance}м {stroke}, {format_time(time_seconds)}, "
            f"{pool}м, {session_type}, {swim_date}"
        )
    await message.answer("\n".join(lines), parse_mode="HTML")


@dp.message(Command("delete"))
async def cmd_delete(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    arg = message.text.partition(" ")[2].strip()
    if not arg:
        await message.answer(
            "Напиши так:\n"
            "/delete 15 — удалить одну запись\n"
            "/delete 15 16 20 — удалить несколько по id\n"
            "/delete Аня - 200 - брасс - 4:00.00 - сорев — удалить по содержимому"
        )
        return

    id_tokens = arg.replace(",", " ").split()
    if id_tokens and all(t.isdigit() for t in id_tokens):
        ids = [int(t) for t in id_tokens]
        deleted = sum(1 for i in ids if delete_result(i))
        if deleted:
            reindex_results()
        await message.answer(
            f"Удалено записей: {deleted} из {len(ids)}."
            + (" Номера записей пересчитаны." if deleted else "")
        )
        return

    parsed = parse_basic(arg)
    if not parsed:
        await message.answer("Не понял, что удалять. Формат: /delete Аня - 200 - брасс - 4:00.00")
        return
    name, distance, stroke, time_seconds, session_type = parsed
    matches = find_matching_results(name, distance, stroke, time_seconds, session_type)
    if not matches:
        await message.answer("Не нашёл такую запись.")
        return
    if len(matches) == 1:
        delete_result(matches[0][0])
        reindex_results()
        await message.answer("Удалено ✅ Номера записей пересчитаны.")
        return

    lines = ["Нашёл несколько подходящих записей, уточни через /delete <id>:"]
    for rid, st, swim_date, pool in matches:
        lines.append(f"#{rid} — {st}, бассейн {pool}м, {swim_date}")
    await message.answer("\n".join(lines))


@dp.message(Command("edit"))
async def cmd_edit(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    rest = message.text.partition(" ")[2].strip()
    id_part, _, record_part = rest.partition(" ")
    if not id_part.isdigit() or not record_part.strip():
        await message.answer(
            "Напиши так: /edit 15 - Аня - 200 - брасс - 3:58.00 - сорев\n"
            "(id смотри через /last)"
        )
        return
    record, warnings = parse_record_full(record_part)
    if not record:
        await message.answer("Не смог разобрать новую запись, проверь формат.")
        return
    ok = update_result(int(id_part), record)
    if ok:
        text = f"Исправлено ✅ #{id_part}: " + build_confirmation(record, warnings, None, None).split("\n", 1)[0]
        await message.answer(text, parse_mode="HTML")
    else:
        await message.answer("Запись с таким id не найдена.")


def build_excel(rows, title="Результаты"):
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    headers = [
        "Пловец", "Дистанция, м", "Бассейн, м", "Стиль", "Время", "∆t к предыдущему",
        "Тип", "Режим", "Часть", "Инвентарь", "Дата",
    ]
    ws.append(headers)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    last_seen = {}  # (name, distance, stroke) -> время предыдущего результата
    for row in rows:
        (name, distance, stroke, time_seconds, session_type, pool, swim_date,
         training_kind, body_part, fins, paddles, snorkel, recorded_at) = row

        key = (name, distance, stroke)
        prev_time = last_seen.get(key)
        delta = time_seconds - prev_time if prev_time is not None else None
        last_seen[key] = time_seconds

        if session_type == "тренировка":
            kind_str, part_str = training_kind or "", body_part or ""
            equip_str = equipment_str(fins, paddles, snorkel)
        else:
            kind_str, part_str, equip_str = "", "", ""

        ws.append([
            name, distance, pool, stroke, format_time(time_seconds), format_delta(delta),
            session_type, kind_str, part_str, equip_str, swim_date or recorded_at.split("T")[0],
        ])

    for col_idx, header in enumerate(headers, start=1):
        max_len = max(
            [len(str(header))] + [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(2, ws.max_row + 1)]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@dp.message(Command("table"))
async def cmd_table(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Общая таблица доступна только тренеру/владельцу. Свой прогресс смотри через /my.")
        return
    arg = message.text.partition(" ")[2].strip()
    name = arg.title() if arg else None
    rows = get_all_results(name)
    if not rows:
        await message.answer("Пока нет данных для таблицы.")
        return
    buf = build_excel(rows, title=name or "Все пловцы")
    filename = f"swim_results_{name or 'all'}.xlsx"
    await message.answer_document(BufferedInputFile(buf.read(), filename=filename))

# ---------------------------------------------------------------------------
# НОРМАТИВЫ (добавляет/удаляет только тренер/владелец, смотреть могут все)
# ---------------------------------------------------------------------------

def parse_one_addnorm(line):
    """Разбирает одну строку норматива вида «Дистанция - Стиль - [Бассейн -] Название - Время».
    Возвращает None для пустой строки, иначе (True, текст об успехе) или (False, текст ошибки)."""
    line = line.strip()
    if not line:
        return None
    # на случай если каждая строка снова начинается с /addnorm или !добавнорм
    line = re.sub(r"^(?:/addnorm|!добавнорм)\s*", "", line, flags=re.I).strip()
    if not line:
        return None

    parts = split_fields(line)
    if len(parts) < 4:
        return False, f"«{line}» — не смог разобрать. Формат: Дистанция - Стиль - [Бассейн -] Название - Время"

    distance_digits = re.sub(r"[^\d]", "", parts[0])
    if not distance_digits:
        return False, f"«{line}» — не понял дистанцию"
    distance = int(distance_digits)
    stroke = normalize_stroke(parts[1])

    rest_parts = parts[2:]
    pool = 25
    if rest_parts and re.match(r"^\d+м?$", rest_parts[0].strip().lower()):
        pool = int(re.sub(r"[^\d]", "", rest_parts[0]))
        rest_parts = rest_parts[1:]

    if len(rest_parts) < 2:
        return False, f"«{line}» — не хватает названия норматива и/или времени"
    *label_parts, time_part = rest_parts
    label = " ".join(label_parts).strip()
    try:
        time_seconds = parse_time(time_part)
    except ValueError:
        return False, f"«{line}» — не смог разобрать время норматива"

    add_standard(distance, stroke, pool, label, time_seconds)
    return True, f"✅ {distance}м {stroke}, бассейн {pool}м, «{label}» — {format_time(time_seconds)}"


@dp.message(Command("addnorm"))
async def cmd_addnorm(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    rest = message.text.partition(" ")[2]
    if not rest.strip():
        await message.answer(
            "Напиши так: /addnorm 100 - брасс - 25 - 3 юн - 1:45.00\n"
            "(Дистанция - Стиль - Бассейн - Название норматива - Время; бассейн можно "
            "не указывать — тогда 25м по умолчанию)\n\n"
            "Можно сразу несколько нормативов — каждый на отдельной строке:\n"
            "<code>/addnorm 50 - вольный стиль - 25 - МСМК - 21.29\n"
            "50 - вольный стиль - 25 - МС - 22.65\n"
            "50 - вольный стиль - 25 - КМС - 23.40</code>",
            parse_mode="HTML",
        )
        return

    added, errors = [], []
    for line in rest.splitlines():
        result = parse_one_addnorm(line)
        if result is None:
            continue
        ok, text = result
        (added if ok else errors).append(text)

    reply_parts = []
    if added:
        reply_parts.append("Добавлено:\n" + "\n".join(added))
    if errors:
        reply_parts.append("Не удалось разобрать:\n" + "\n".join(errors))
    if not reply_parts:
        reply_parts.append("Не удалось разобрать ни одной строки.")
    await message.answer("\n\n".join(reply_parts))


@dp.message(Command("delnorm"))
async def cmd_delnorm(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    arg = message.text.partition(" ")[2].strip()
    if not arg.isdigit():
        await message.answer("Напиши так: /delnorm 5 (id смотри через /normativy)")
        return
    ok = delete_standard(int(arg))
    await message.answer("Удалено ✅" if ok else "Норматив с таким id не найден.")


@dp.message(Command("normativy"))
async def cmd_normativy(message: Message):
    arg = message.text.partition(" ")[2].strip()
    distance, stroke = None, None
    if arg:
        parts = split_fields(arg)
        distance_digits = re.sub(r"[^\d]", "", parts[0]) if parts else ""
        if distance_digits:
            distance = int(distance_digits)
        if len(parts) >= 2:
            stroke = normalize_stroke(parts[1])

    rows = get_standards(distance, stroke)
    if not rows:
        await message.answer("Нормативов пока нет." if not arg else "Нормативов на это не нашлось.")
        return

    await message.answer(format_standards_message(rows), parse_mode="HTML")

# ---------------------------------------------------------------------------
# ТРЕНИРОВКИ: ПЛАН И ПРОСМОТР
# ---------------------------------------------------------------------------

@dp.message(Command("addtraining"))
async def cmd_addtraining(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Добавлять тренировки может только тренер или владелец.")
        return

    rest = message.text.partition(" ")[2].strip()
    if not rest:
        await message.answer("Формат: /addtraining 16.09-20.09")
        return

    # Если после диапазона уже передан текст, сохраняем его на все даты.
    # Для разных тренировок по дням используем мастер: сначала диапазон,
    # затем бот по очереди просит тренировку для каждого дня.
    m = re.match(r"^(\S+?)(?:\s+-\s+(.+))?$", rest, re.S)
    if not m:
        await message.answer("Формат: /addtraining 16.09-20.09")
        return

    dates = parse_training_dates(m.group(1))
    if not dates:
        await message.answer("Не понял даты. Пример: /addtraining 16.09-20.09")
        return

    training_text = (m.group(2) or "").strip()
    if training_text:
        lines = training_text.splitlines()
        title = "Тренировка"
        description = training_text
        if lines and re.match(r"^Название\s*:", lines[0], re.I):
            title = re.sub(r"^Название\s*:\s*", "", lines[0], flags=re.I).strip() or title
            description = "\n".join(lines[1:]).strip() or title
        for d in dates:
            save_training(d, title, description)
        await message.answer(f"Тренировка сохранена на {len(dates)} дней ✅")
        return

    PENDING_TRAININGS[message.from_user.id] = {"dates": dates, "index": 0}
    d = dates[0]
    label = datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y")
    await message.answer(
        f"📅 Диапазон создан: {len(dates)} дней.\n\n"
        f"Теперь введи <b>отдельную тренировку</b> для {label}.\n"
        "После отправки бот перейдёт к следующему дню.\n\n"
        "Можно написать, например:\n"
        "<b>Название: Силовая</b>\n50м ноги\n8×100 кроль\n200м заминка",
        parse_mode="HTML"
    )


@dp.callback_query(F.data.startswith("plan_day:"))
async def on_plan_day(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Только для тренера.", show_alert=True)
        return
    _, date_str = callback.data.split(":", 1)
    state = PENDING_TRAININGS.get(callback.from_user.id)
    if not state or date_str not in state["dates"]:
        await callback.answer("Этот план уже закрыт.", show_alert=True)
        return
    state["index"] = state["dates"].index(date_str)
    label = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")
    await callback.message.answer(f"Введи отдельную тренировку на <b>{label}</b>.", parse_mode="HTML")
    await callback.answer()


@dp.callback_query(F.data == "plan_cancel")
async def on_plan_cancel(callback: CallbackQuery):
    PENDING_TRAININGS.pop(callback.from_user.id, None)
    await callback.message.answer("Создание тренировок отменено.")
    await callback.answer()


async def send_today_training(message: Message):
    training = get_training(datetime.now().strftime("%Y-%m-%d"))
    if not training:
        builder = InlineKeyboardBuilder()
        builder.button(text="⬅️ Предыдущая тренировка", callback_data="view_prev_training")
        builder.button(text="📐 Нормативы", callback_data="nrm_open")
        builder.adjust(1)
        await message.answer("Сегодня выходной 🛌", reply_markup=builder.as_markup())
        return
    builder = InlineKeyboardBuilder()
    builder.button(text="⬅️ Предыдущая тренировка", callback_data="view_prev_training")
    builder.button(text="📐 Нормативы", callback_data="nrm_open")
    builder.adjust(1)
    await message.answer(format_training(training), parse_mode="HTML", reply_markup=builder.as_markup())


async def send_prev_training(message: Message):
    today = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT id, training_date, title, description FROM trainings "
        "WHERE training_date < ? ORDER BY training_date DESC LIMIT 1", (today,)
    ).fetchone()
    conn.close()
    if not row:
        await message.answer("Предыдущих тренировок пока нет.")
        return
    builder = InlineKeyboardBuilder()
    builder.button(text="📅 Тренировка на сегодня", callback_data="view_today_training")
    builder.adjust(1)
    await message.answer(format_training(row), parse_mode="HTML", reply_markup=builder.as_markup())


@dp.callback_query(F.data == "view_today_training")
async def on_view_today_training(callback: CallbackQuery):
    await send_today_training(callback.message)
    await callback.answer()


@dp.callback_query(F.data == "view_prev_training")
async def on_view_prev_training(callback: CallbackQuery):
    await send_prev_training(callback.message)
    await callback.answer()


# ---------------------------------------------------------------------------
# ПРОСМОТР НОРМАТИВОВ ЧЕРЕЗ КНОПКИ (доступно всем)
# ---------------------------------------------------------------------------

@dp.callback_query(F.data == "nrm_open")
async def on_nrm_open(callback: CallbackQuery):
    builder = InlineKeyboardBuilder()
    builder.button(text="25м бассейн", callback_data="nrm_pool:25")
    builder.button(text="50м бассейн", callback_data="nrm_pool:50")
    builder.adjust(2)
    await callback.message.answer("Какой бассейн?", reply_markup=builder.as_markup())
    await callback.answer()


@dp.callback_query(F.data.startswith("nrm_pool:"))
async def on_nrm_pool(callback: CallbackQuery):
    pool = int(callback.data.split(":", 1)[1])
    builder = InlineKeyboardBuilder()
    builder.button(text="📋 Все нормативы", callback_data=f"nrm_all:{pool}")
    builder.button(text="🎯 Определённая дистанция", callback_data=f"nrm_distpick:{pool}")
    builder.adjust(1)
    await callback.message.answer(f"Бассейн {pool}м. Что показать?", reply_markup=builder.as_markup())
    await callback.answer()


@dp.callback_query(F.data.startswith("nrm_all:"))
async def on_nrm_all(callback: CallbackQuery):
    pool = int(callback.data.split(":", 1)[1])
    rows = get_standards(pool=pool)
    if not rows:
        await callback.message.answer(f"Нормативов для бассейна {pool}м пока нет.")
    else:
        await callback.message.answer(format_standards_message(rows), parse_mode="HTML")
    await callback.answer()


@dp.callback_query(F.data.startswith("nrm_distpick:"))
async def on_nrm_distpick(callback: CallbackQuery):
    pool = int(callback.data.split(":", 1)[1])
    distances = get_standard_distances(pool)
    if not distances:
        await callback.message.answer(f"Нормативов для бассейна {pool}м пока нет.")
        await callback.answer()
        return
    builder = InlineKeyboardBuilder()
    for d in distances:
        builder.button(text=f"{d}м", callback_data=f"nrm_dist:{pool}:{d}")
    builder.adjust(3)
    await callback.message.answer("Выбери дистанцию:", reply_markup=builder.as_markup())
    await callback.answer()


@dp.callback_query(F.data.startswith("nrm_dist:"))
async def on_nrm_dist(callback: CallbackQuery):
    _, pool_str, distance_str = callback.data.split(":", 2)
    pool, distance = int(pool_str), int(distance_str)
    strokes = get_standard_strokes(pool, distance)
    builder = InlineKeyboardBuilder()
    for s in strokes:
        builder.button(text=s.capitalize(), callback_data=f"nrm_final:{pool}:{distance}:{s}")
    builder.button(text="Пропустить (все стили)", callback_data=f"nrm_final:{pool}:{distance}:_all")
    builder.adjust(2)
    await callback.message.answer("Указать стиль или пропустить этот шаг?", reply_markup=builder.as_markup())
    await callback.answer()


@dp.callback_query(F.data.startswith("nrm_final:"))
async def on_nrm_final(callback: CallbackQuery):
    _, pool_str, distance_str, stroke = callback.data.split(":", 3)
    pool, distance = int(pool_str), int(distance_str)
    stroke_filter = None if stroke == "_all" else stroke
    rows = get_standards(distance=distance, stroke=stroke_filter, pool=pool)
    if not rows:
        await callback.message.answer("Нормативов на это не нашлось.")
    else:
        await callback.message.answer(format_standards_message(rows), parse_mode="HTML")
    await callback.answer()


@dp.callback_query(F.data == "view_my_results")
async def on_view_my_results(callback: CallbackQuery):
    await cmd_my(callback.message)
    await callback.answer()


@dp.message(Command("training"))
async def cmd_training(message: Message):
    await send_today_training(message)


@dp.message(Command("prevtraining"))
async def cmd_prevtraining(message: Message):
    await send_prev_training(message)


@dp.message(F.text.regexp(r"^!треня$"))
async def cmd_ru_training(message: Message):
    await send_today_training(message)


@dp.message(F.text.regexp(r"^!предтреня$"))
async def cmd_ru_prevtraining(message: Message):
    await send_prev_training(message)


@dp.message(Command("trainings"))
async def cmd_trainings(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Эта команда только для тренера/владельца.")
        return
    rows = get_training_dates()
    if not rows:
        await message.answer("Запланированных тренировок нет.")
        return
    lines = ["<b>Запланированные тренировки:</b>"]
    for d, title in rows:
        lines.append(f"• {datetime.strptime(d, '%Y-%m-%d').strftime('%d.%m.%Y')} — {title}")
    await message.answer("\n".join(lines), parse_mode="HTML")


@dp.message(Command("deltraining"))
async def cmd_deltraining(message: Message):
    if not is_admin(message.from_user.id):
        await message.answer("Удалять тренировки может только тренер или владелец.")
        return
    d = parse_date_value(message.text.partition(" ")[2].strip())
    if not d:
        await message.answer("Формат: /deltraining 18.09")
        return
    await message.answer("Тренировка удалена ✅" if delete_training(d) else "На эту дату тренировки нет.")


# ---------------------------------------------------------------------------
# ПРОСМОТР СВОЕГО ПРОГРЕССА (доступно всем)
# ---------------------------------------------------------------------------

@dp.message(Command("my"))
async def cmd_my(message: Message):
    uid = message.from_user.id
    name = get_linked_name(uid)
    if not name:
        await message.answer("Сначала привяжи имя: <code>/iam Твоё Имя</code>", parse_mode="HTML")
        return
    combos = get_distances_for_name(name)
    if not combos:
        await message.answer(f"Пока нет ни одного результата для {name}.")
        return
    builder = InlineKeyboardBuilder()
    for distance, stroke in combos:
        label = f"{distance}м {stroke}"
        builder.button(text=label, callback_data=f"prog:{distance}:{stroke}:{uid}")
    builder.adjust(2)
    await message.answer(f"Выбери дистанцию, {name}:", reply_markup=builder.as_markup())


@dp.callback_query(F.data.startswith("prog:"))
async def on_progress_click(callback: CallbackQuery):
    _, distance_str, stroke, owner_id_str = callback.data.split(":", 3)
    distance = int(distance_str)
    owner_id = int(owner_id_str)

    if callback.from_user.id != owner_id:
        await callback.answer("Это не твои кнопки 🙂", show_alert=True)
        return

    name = get_linked_name(owner_id)
    if not name:
        await callback.answer("Имя не найдено.", show_alert=True)
        return

    best_comp_row = get_best_with_pool(name, distance, stroke, "соревнования")
    best_train_row = get_best_with_pool(name, distance, stroke, "тренировка")
    best_comp = best_comp_row[0] if best_comp_row else None
    best_train = best_train_row[0] if best_train_row else None
    last_comp = get_entries(name, distance, stroke, "соревнования", limit=2)
    last_train = get_entries(name, distance, stroke, "тренировка", limit=2)

    lines = [f"📊 <b>{name} — {distance}м {stroke}</b>\n"]

    lines.append(f"🏆 Лучший на соревнованиях: {format_time(best_comp) if best_comp else '—'}")
    lines.append(f"🏋️ Лучший на тренировках: {format_time(best_train) if best_train else '—'}\n")

    best_row = best_comp_row or best_train_row
    if best_row:
        status = get_norm_status(distance, stroke, best_row[1], best_row[0])
        if status:
            if status["achieved_all"]:
                lines.append(f"🥇 По лучшему результату выполнены все нормативы (лучший: {status['top_label']})\n")
            else:
                lines.append(
                    f"📐 До норматива «{status['label']}» ({format_time(status['target_time'])}) "
                    f"не хватает {format_time(status['gap'])}\n"
                )

    lines.append("<b>Соревнования (последние):</b>")
    if last_comp:
        for t, dt in last_comp:
            lines.append(f"• {format_time(t)} ({dt})")
        if len(last_comp) == 2:
            delta = last_comp[0][0] - last_comp[1][0]
            arrow = "⬇️ быстрее" if delta < 0 else ("⬆️ медленнее" if delta > 0 else "➖ так же")
            lines.append(f"Изменение: {arrow} на {format_time(abs(delta))}")
    else:
        lines.append("• пока нет результатов")

    lines.append("\n<b>Тренировки (последние):</b>")
    if last_train:
        for t, dt in last_train:
            lines.append(f"• {format_time(t)} ({dt})")
        if len(last_train) == 2:
            delta = last_train[0][0] - last_train[1][0]
            arrow = "⬇️ быстрее" if delta < 0 else ("⬆️ медленнее" if delta > 0 else "➖ так же")
            lines.append(f"Изменение: {arrow} на {format_time(abs(delta))}")
    else:
        lines.append("• пока нет результатов")

    await callback.message.answer("\n".join(lines), parse_mode="HTML")
    await callback.answer()

# ---------------------------------------------------------------------------
# АВТОРАСПОЗНАВАНИЕ СООБЩЕНИЙ (только тренер/владелец)
# ---------------------------------------------------------------------------

@dp.message(F.text.regexp(r"^!результат$"))
async def cmd_ru_my(message: Message):
    await cmd_my(message)


# ---------------------------------------------------------------------------
# РУССКИЕ АЛИАСЫ ДЛЯ ОСТАЛЬНЫХ КОМАНД
# ---------------------------------------------------------------------------
# Каждый обработчик просто вызывает уже существующую функцию — вся разборка
# аргументов внутри неё делается через message.text.partition(" ")[2], поэтому
# ей всё равно, каким словом (/команда или !алиас) начинается сообщение.

@dp.message(F.text.regexp(r"^!помощь$"))
async def cmd_ru_help(message: Message):
    await cmd_help(message)


@dp.message(F.text.regexp(r"^!ктоя$"))
async def cmd_ru_whoami(message: Message):
    await cmd_whoami(message)


@dp.message(F.text.regexp(r"^!яэто(?:\s|$)"))
async def cmd_ru_iam(message: Message):
    await cmd_iam(message)


@dp.message(F.text.regexp(r"^!привязать(?:\s|$)"))
async def cmd_ru_link(message: Message):
    await cmd_link(message)


@dp.message(F.text.regexp(r"^!отвязать(?:\s|$)"))
async def cmd_ru_unlink(message: Message):
    await cmd_unlink(message)


@dp.message(F.text.regexp(r"^!пользователи$"))
async def cmd_ru_users(message: Message):
    await cmd_users(message)


@dp.message(F.text.regexp(r"^!добавить(?:\s|$)"))
async def cmd_ru_add(message: Message):
    await cmd_add(message)


@dp.message(F.text.regexp(r"^!последние$"))
async def cmd_ru_last(message: Message):
    await cmd_last(message)


@dp.message(F.text.regexp(r"^!удалить(?:\s|$)"))
async def cmd_ru_delete(message: Message):
    await cmd_delete(message)


@dp.message(F.text.regexp(r"^!редакт(?:\s|$)"))
async def cmd_ru_edit(message: Message):
    await cmd_edit(message)


@dp.message(F.text.regexp(r"^!таблица(?:\s|$)"))
async def cmd_ru_table(message: Message):
    await cmd_table(message)


@dp.message(F.text.regexp(r"^!добавнорм(?:\s|$)"))
async def cmd_ru_addnorm(message: Message):
    await cmd_addnorm(message)


@dp.message(F.text.regexp(r"^!удалнорм(?:\s|$)"))
async def cmd_ru_delnorm(message: Message):
    await cmd_delnorm(message)


@dp.message(F.text.regexp(r"^!нормативы(?:\s|$)"))
async def cmd_ru_normativy(message: Message):
    await cmd_normativy(message)


@dp.message(F.text.regexp(r"^!добавтреню(?:\s|$)"))
async def cmd_ru_addtraining(message: Message):
    await cmd_addtraining(message)


@dp.message(F.text.regexp(r"^!тренировки$"))
async def cmd_ru_trainings(message: Message):
    await cmd_trainings(message)


@dp.message(F.text.regexp(r"^!удалтреню(?:\s|$)"))
async def cmd_ru_deltraining(message: Message):
    await cmd_deltraining(message)


@dp.message(F.text)
async def catch_all(message: Message):
    text = (message.text or "").strip()

    # Русские команды для просмотра результатов/тренировок обрабатываются
    # отдельными хендлерами выше.
    if text.startswith("/") or text.startswith("!"):
        return

    # Заполнение диапазона тренировок тренером: каждое сообщение — отдельный день.
    if is_admin(message.from_user.id):
        state = PENDING_TRAININGS.get(message.from_user.id)
        if state:
            d = state["dates"][state["index"]]
            lines = text.splitlines()
            title = "Тренировка"
            description = text
            if lines and re.match(r"^Название\s*:", lines[0], re.I):
                title = re.sub(r"^Название\s*:\s*", "", lines[0], flags=re.I).strip() or title
                description = "\n".join(lines[1:]).strip() or title
            save_training(d, title, description)

            state["index"] += 1
            if state["index"] >= len(state["dates"]):
                PENDING_TRAININGS.pop(message.from_user.id, None)
                await message.answer("Все тренировки на диапазон сохранены ✅")
                return

            next_d = state["dates"][state["index"]]
            label = datetime.strptime(next_d, "%Y-%m-%d").strftime("%d.%m.%Y")
            await message.answer(f"✅ Сохранено. Теперь тренировка на <b>{label}</b>.", parse_mode="HTML")
            return

        # Старое автораспознавание результатов.
        record, _ = parse_record_full(text)
        if not record and not parse_batch_records(text):
            return
        await handle_incoming_text(message, text)


async def main():
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
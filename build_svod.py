# -*- coding: utf-8 -*-
"""
Сборщик сводного графика ремонтов ЛЭП и сетевого оборудования.

Принцип работы
--------------
Скрипт ищет в своей папке файлы:
    * Проект Арх РДУ.xlsx                                   — экспорт из ПК «Ремонты»
    * Проект Коми РДУ.xlsx                                  — экспорт из ПК «Ремонты»
    * Приоритет строк по группам для сводного графика.xlsx — справочник приоритетов

Если каких-то файлов в корне нет — пытается их найти в подпапке «Исходные материалы».

На выходе в той же корневой папке появляется файл:
    Сводный график ремонтов ЛЭП и сетевого оборудования на <месяц> <год> г.xlsx

Запуск
------
    python build_svod.py                    — собрать; год/месяц определяются автоматически
    python build_svod.py --year 2026        — указать год вручную
    python build_svod.py --no-normalize     — без текстовой нормализации
    python build_svod.py --collapse-preamble— дополнительно сворачивать преамбулы
                                              «Вывод в ремонт … для проведения …»
    python build_svod.py --dry-run          — ничего не сохранять, только отчёт
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import Counter, defaultdict, OrderedDict
from copy import copy as _copy
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.worksheet import Worksheet

# ------------------------------------------------------------------ КОНСТАНТЫ -

ROOT = Path(__file__).resolve().parent

FALLBACK_DIR = ROOT / "Исходные материалы"

FILE_ARKH = "Проект Арх РДУ.xlsx"
FILE_KOMI = "Проект Коми РДУ.xlsx"
FILE_PRIO = "Приоритет строк по группам для сводного графика.xlsx"

# Итоговое число колонок таблицы (A..Y).
TABLE_COLS = 25
LAST_COL_LETTER = get_column_letter(TABLE_COLS)

# Группы в порядке вывода.
GROUP_OGR = "OGR"          # Ограничения ОЗ
GROUP_LEP220 = "LEP220"    # ЛЭП 220 кВ
GROUP_PS220 = "PS220"      # ПС 220 кВ
GROUP_LEP110 = "LEP110"    # ЛЭП 110 кВ
GROUP_PS110 = "PS110"      # ПС 110 кВ
GROUP_ES = "ES"            # Электростанции
GROUP_ACHR = "ACHR"        # АЧР
GROUP_OTHER = "OTHER"      # Прочее (попадает всё, что не удалось классифицировать)

GROUP_ORDER = [
    GROUP_OGR,
    GROUP_LEP220,
    GROUP_PS220,
    GROUP_LEP110,
    GROUP_PS110,
    GROUP_ES,
    GROUP_ACHR,
    GROUP_OTHER,
]

GROUP_LABELS = {
    GROUP_OGR:    "Ограничения ОЗ",
    GROUP_LEP220: "ЛЭП 220 кВ",
    GROUP_PS220:  "ПС 220 кВ",
    GROUP_LEP110: "ЛЭП 110 кВ",
    GROUP_PS110:  "ПС 110 кВ",
    GROUP_ES:     "Электростанции",
    GROUP_ACHR:   "АЧР",
    GROUP_OTHER:  "Прочее (не классифицировано)",
}

RU_MONTHS_NOM = [
    "",  # dummy для 1-based
    "январь", "февраль", "март",     "апрель", "май",    "июнь",
    "июль",   "август",  "сентябрь", "октябрь", "ноябрь", "декабрь",
]

RU_MONTHS_SHORT = [
    "",
    "янв", "фев", "мар", "апр", "май", "июн",
    "июл", "авг", "сен", "окт", "ноя", "дек",
]

# ---- Высоты строк (пт) ----
ROW_HEIGHT_SECTION = 22.0       # заголовок группы (ЛЭП 220 кВ, …)
ROW_HEIGHT_SUBSECTION = 18.0    # подзаголовок объекта (ПС 220 кВ Вельск, …)
ROW_HEIGHT_TOC = 18.0           # строка оглавления

# ---- Заливка листа «Диаграмма» по виду ремонта (RGB без #) ----
GANTT_COLORS: dict[str, str] = {
    "ТР":  "B6D7A8",   # светло-зелёный
    "СР":  "FFE599",   # светло-жёлтый
    "КР":  "EA9999",   # розово-красный
    "ВПр": "9FC5E8",   # голубой
    "ИСП": "F9CB9C",   # оранжевый
    "ЗРР": "B4A7D6",   # сиреневый
    "БВР": "CCCCCC",   # серый
}
GANTT_COLOR_OTHER = "EEEEEE"
GANTT_COLOR_WEEKEND = "F2F2F2"
GANTT_SHEET_NAME = "Диаграмма"

BACKUP_DIR = ROOT / "backups"
SVOD_FILE_PREFIX = "Сводный график ремонтов"


# ------------------------------------------------------------------ УТИЛИТЫ --

def find_file(name: str) -> Path:
    """Ищет файл в корне, затем в 'Исходные материалы'. Возвращает Path или None."""
    for base in (ROOT, FALLBACK_DIR):
        p = base / name
        if p.exists():
            return p
    return None


def parse_day_month(value, default_year: int) -> tuple[int, int, int] | None:
    """Разбирает '12.05.' / '12.05' / '12.05.2026' / datetime в (год, месяц, день).
    Возвращает None, если распарсить не удалось."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return (value.year, value.month, value.day)
    s = str(value).strip().rstrip(".").strip()
    if not s:
        return None
    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?$", s)
    if not m:
        return None
    day, mon = int(m.group(1)), int(m.group(2))
    year = int(m.group(3)) if m.group(3) else default_year
    if year < 100:
        year += 2000
    return (year, mon, day)


def _cell_text_with_merges(ws: Worksheet, row: int, col: int) -> str:
    """Возвращает текст ячейки; если ячейка внутри объединения и сама пустая —
    вернёт текст «владельца» объединения (top-left-ячейки)."""
    v = ws.cell(row, col).value
    if v is not None and str(v).strip() != "":
        return str(v)
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            owner = ws.cell(mr.min_row, mr.min_col).value
            if owner is not None:
                return str(owner)
            break
    return ""


def validate_project_template(ws: Worksheet, filename: str) -> None:
    """Проверяет, что лист похож на экспорт ПК «Ремонты». Падает с понятным
    сообщением, если формат не распознан."""
    errors: list[str] = []

    # 1. Имя листа.
    if ws.title != "Page1":
        errors.append(
            f"ожидается лист с именем «Page1», найден «{ws.title}»"
        )

    # 2. Шапка и таблица: A6 — «Наименование оборудования», F6/G6 — «Дата ...»,
    #    N6 — «Вид ремонта». Допускаем расхождения в пробелах/регистре.
    def hdr(col: int) -> str:
        return re.sub(r"\s+", " ", _cell_text_with_merges(ws, 6, col)).strip().lower()

    a = hdr(1)
    f = hdr(6)
    g = hdr(7)
    n = hdr(14)

    if "наименован" not in a:
        errors.append("в ячейке A6 не найдено «Наименование оборудования»")
    if not ("дата" in f or "начал" in f or "начал" in g):
        errors.append("в колонках F6/G6 не найдены «Дата начала/окончания»")
    if not ("вид" in n or "ремонт" in n):
        errors.append("в ячейке N6 не найдено «Вид ремонта»")

    if errors:
        print()
        print(f"ОШИБКА: файл «{filename}» выглядит не как экспорт ПК «Ремонты».")
        for e in errors:
            print(f"  • {e}")
        print()
        print("Ожидаемый формат: лист «Page1», шапка в строках 1–6, заголовок")
        print("  таблицы: A «Наименование оборудования», F/G «Дата начала/"
              "окончания», N «Вид ремонта».")
        print("Если ПК «Ремонты» выдал новый формат — сообщите разработчику, "
              "приложив файл.")
        sys.exit(3)


def month_day_count(year: int, month: int) -> int:
    """Количество дней в указанном месяце."""
    if month == 12:
        nxt = datetime(year + 1, 1, 1)
    else:
        nxt = datetime(year, month + 1, 1)
    return (nxt - datetime(year, month, 1)).days


def copy_cell_style(src, dst):
    """Копирует стиль исходной ячейки в целевую (возможно, из другой книги)."""
    if src.has_style:
        dst.font = _copy(src.font)
        dst.fill = _copy(src.fill)
        dst.border = _copy(src.border)
        dst.alignment = _copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = _copy(src.protection)


def copy_cell(src, dst):
    dst.value = src.value
    copy_cell_style(src, dst)


def copy_row_full(src_ws: Worksheet, src_row: int,
                  dst_ws: Worksheet, dst_row: int,
                  ncols: int = TABLE_COLS):
    for c in range(1, ncols + 1):
        copy_cell(src_ws.cell(src_row, c), dst_ws.cell(dst_row, c))
    rh = src_ws.row_dimensions[src_row].height
    if rh is not None:
        dst_ws.row_dimensions[dst_row].height = rh


def copy_merges_in_row(src_ws: Worksheet, src_row: int,
                       dst_ws: Worksheet, dst_row: int,
                       ncols: int = TABLE_COLS):
    """Копирует объединения, находящиеся в указанной строке источника."""
    ranges = list(src_ws.merged_cells.ranges)
    for mr in ranges:
        if mr.min_row != src_row or mr.max_row != src_row:
            continue
        if mr.min_col > ncols:
            continue
        lo = mr.min_col
        hi = min(mr.max_col, ncols)
        rng = f"{get_column_letter(lo)}{dst_row}:{get_column_letter(hi)}{dst_row}"
        try:
            dst_ws.merge_cells(rng)
        except Exception:
            pass


def copy_column_widths(src_ws: Worksheet, dst_ws: Worksheet,
                       ncols: int = TABLE_COLS + 1):
    for c in range(1, ncols + 1):
        letter = get_column_letter(c)
        w = src_ws.column_dimensions[letter].width
        if w:
            dst_ws.column_dimensions[letter].width = w


# -------------------------------------------------------- ПАРСИНГ ПРОЕКТА ----

def is_section_row(ws: Worksheet, row: int, ncols: int = TABLE_COLS) -> bool:
    """Строка-подзаголовок: объединена на всю ширину (A..Y)."""
    for mr in ws.merged_cells.ranges:
        if (mr.min_row == row and mr.max_row == row
                and mr.min_col == 1 and mr.max_col >= ncols):
            return True
    return False


def is_equipment_row(ws: Worksheet, row: int) -> bool:
    """Строка оборудования: в A что-то есть, при этом это не секция и не
    строка подписи (у подписей A пусто)."""
    a = ws.cell(row, 1).value
    if a is None or str(a).strip() == "":
        return False
    return not is_section_row(ws, row)


def find_data_bounds(ws: Worksheet, ncols: int = TABLE_COLS) -> tuple[int, int, int]:
    """Возвращает (header_last_row, data_last_row, signatures_start_row).

    Правила:
      * Шапка таблицы занимает строки 1..6 — это константа формата проекта.
      * Данные — строки с непустым A (секции или оборудование).
      * Подписи — начинаются после последней «data»-строки, могут содержать
        пустые промежутки между подписывающими лицами.
    """
    header_last = 6
    last_data_row = header_last
    for r in range(header_last + 1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is not None and str(a).strip() != "":
            last_data_row = r

    # sig_start — первая непустая строка после data-блока.
    sig_start = last_data_row + 1
    while sig_start <= ws.max_row:
        row_empty = True
        for c in range(1, ncols + 1):
            v = ws.cell(sig_start, c).value
            if v is not None and str(v).strip() != "":
                row_empty = False
                break
        if row_empty:
            sig_start += 1
        else:
            break

    data_last = last_data_row
    return header_last, data_last, sig_start


def extract_records(ws: Worksheet, rdu: str, default_year: int,
                    src_key: str) -> list[dict]:
    """Возвращает список записей с исходных строк оборудования.

    Каждая запись содержит ссылку на исходный лист и номер строки — это
    позволит затем скопировать её «как есть» (со всеми стилями и объединениями).
    """
    header_last, data_last, sig_start = find_data_bounds(ws)
    recs: list[dict] = []
    current_section = None
    for r in range(header_last + 1, data_last + 1):
        a = ws.cell(r, 1).value
        if a is None or (isinstance(a, str) and a.strip() == ""):
            continue
        name = str(a).strip()
        if is_section_row(ws, r):
            current_section = name
            continue
        # строка оборудования
        start_raw = ws.cell(r, 6).value
        end_raw   = ws.cell(r, 7).value
        start = parse_day_month(start_raw, default_year)
        end   = parse_day_month(end_raw,   default_year)
        recs.append({
            "rdu":     rdu,                        # 'Арх' / 'Коми'
            "section": current_section or "",      # подзаголовок проекта
            "name":    name,                       # значение в столбце A
            "start":   start,                      # (y, m, d) или None
            "end":     end,                        # (y, m, d) или None
            "src_ws":  ws,
            "src_row": r,
            "src_key": src_key,                    # 'arkh' / 'komi' для отладки
        })
    return recs


# -------------------------------------------------- СПРАВОЧНИК ПРИОРИТЕТОВ ---

def load_priority(path: Path) -> dict:
    """Возвращает словарь с порядками объектов по группам.

    Правило разбора:
      * Заголовок раздела — строка, заканчивающаяся двоеточием
        (например «ПС 220 кВ:», «Электростанции:», «АЧР:»).
        Подсказки вида «сначала ПС 220 кВ ОЗ Архангельского РДУ:» тоже
        оканчиваются ":", но идентификатор раздела по ним не меняется.
      * Служебная строка «отсортировать даты начала…» — игнорируется.
      * Элементы вида «Ограничения ОЗ Архангельского РДУ» сами задают
        раздел OGR (в справочнике у этой группы нет отдельного заголовка).
      * Прочие строки — элементы списка текущего раздела.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    def section_of(low: str) -> str | None:
        if "ограничения оз" in low:            return "OGR"
        if low.startswith("лэп 220"):          return "LEP220"
        if low.startswith("пс 220"):           return "PS220"
        if low.startswith("лэп 110"):          return "LEP110"
        if low.startswith("пс 110"):           return "PS110"
        if low.startswith("электростанц"):     return "ES"
        if low.startswith("ачр"):              return "ACHR"
        return None

    current: str | None = None
    data: dict[str, list[str]] = defaultdict(list)

    for r in range(1, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if b is None:
            continue
        text = str(b).strip()
        if text == "" or text.startswith("Приоритет"):
            continue

        low = text.lower().rstrip(":").strip()

        if "отсортировать" in low:
            continue

        if text.rstrip().endswith(":"):
            # Заголовок раздела или внутренняя подсказка.
            sec = section_of(low)
            if sec:
                current = sec
            continue

        # Элементы списка «Ограничения ОЗ …» сами открывают раздел OGR
        # (отдельного заголовка группы в справочнике нет).
        if low.startswith("ограничения оз"):
            current = "OGR"
            data["OGR"].append(text)
            continue

        if current is not None:
            data[current].append(text)

    return {
        "OGR":    data.get("OGR", []),
        "LEP220": data.get("LEP220", []),
        "PS220":  data.get("PS220", []),
        "LEP110": data.get("LEP110", []),
        "PS110":  data.get("PS110", []),
        "ES":     data.get("ES", []),
        "ACHR":   data.get("ACHR", []),
    }


# ------------------------------------------------------------ КЛАССИФИКАЦИЯ --

RE_ACHR       = re.compile(r"(?i)(?:снижение объ[её]ма нагрузки|ачр)")
RE_OGRAN      = re.compile(r"(?i)ограничени\w*\s+режим")
RE_LINE       = re.compile(r"(?i)^\s*вл\s")
RE_220        = re.compile(r"(?i)220\s*кв")
RE_110        = re.compile(r"(?i)110\s*кв")
RE_PS_SECT    = re.compile(r"(?i)^\s*пс\s+(220|110)\s*кв")
RE_ES_SECT    = re.compile(r"(?i)(тэц|грэс)")  # 'ТЭЦ СЛПК', 'Сосногорская ТЭЦ', 'Печорская ГРЭС'

def classify(rec: dict) -> tuple[str, str]:
    """Возвращает (group_key, subgroup_label).
    subgroup_label — название ПС/Электростанции/ОЗ для групп, где это уместно,
    либо "" для «плоских» групп (ЛЭП, АЧР, Прочее)."""

    name = rec["name"] or ""
    section = rec["section"] or ""

    if RE_OGRAN.search(name):
        sub = f"Ограничения ОЗ {rec['rdu']} РДУ"
        return (GROUP_OGR, sub)

    if RE_ACHR.search(name):
        return (GROUP_ACHR, "")

    if RE_LINE.match(name):
        if RE_220.search(name):
            return (GROUP_LEP220, "")
        if RE_110.search(name):
            return (GROUP_LEP110, "")
        # ВЛ без явной отметки кВ — попробуем по секции
        if RE_220.search(section):
            return (GROUP_LEP220, "")
        if RE_110.search(section):
            return (GROUP_LEP110, "")
        return (GROUP_OTHER, "")

    # электростанция — определяем по секции
    if RE_ES_SECT.search(section):
        return (GROUP_ES, section.strip())

    m = RE_PS_SECT.match(section)
    if m:
        kv = m.group(1)
        if kv == "220":
            return (GROUP_PS220, section.strip())
        if kv == "110":
            return (GROUP_PS110, section.strip())

    return (GROUP_OTHER, section.strip())


# -------------------------------- ГРУППИРОВКА И СОРТИРОВКА РЕЗУЛЬТАТОВ ------

def _norm(s: str) -> str:
    """Нормализация названия объекта для сопоставления со справочником:
    удаляет лишние пробелы и кавычки-варианты, приводит к нижнему регистру."""
    s = s or ""
    s = s.replace("«", "").replace("»", "").replace('"', "").replace("'", "")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def subgroup_index(priority_list: list[str], label: str) -> int:
    """Возвращает индекс позиции объекта в справочнике, либо большое число,
    если объект не найден (такие уходят в конец группы)."""
    key = _norm(label)
    for i, item in enumerate(priority_list):
        if _norm(item) == key:
            return i
    return 10_000  # не найдено — в конец


def start_sort_key(rec: dict) -> tuple:
    s = rec.get("start")
    if s is None:
        # записи без даты — в самый конец своей группы
        return (9999, 99, 99)
    return s


def group_and_sort(records: list[dict], priority: dict) -> dict:
    """Возвращает OrderedDict: group_key -> list[record] (уже в порядке вывода).

    Для групп с подгруппами (ПС/Электростанции/Ограничения) записи внутри
    одной подгруппы идут подряд; порядок подгрупп задаётся справочником."""
    buckets = defaultdict(list)
    for rec in records:
        g, sub = classify(rec)
        rec["group"] = g
        rec["subgroup"] = sub
        buckets[g].append(rec)

    ordered: "OrderedDict[str, list[dict]]" = OrderedDict()
    unknown_warnings: list[str] = []

    for g in GROUP_ORDER:
        if g not in buckets or not buckets[g]:
            continue

        items = buckets[g]

        if g in (GROUP_LEP220, GROUP_LEP110, GROUP_ACHR, GROUP_OTHER):
            items.sort(key=start_sort_key)
        elif g == GROUP_OGR:
            # Сначала Арх, потом Коми; внутри — по дате.
            def ogr_key(r):
                rdu_order = 0 if r["rdu"] == "Арх" else 1
                return (rdu_order, ) + tuple(start_sort_key(r))
            items.sort(key=ogr_key)
        else:
            # PS220 / PS110 / ES — по справочнику, внутри подгруппы — по дате
            plist = priority.get(g, [])
            def sort_key(r):
                idx = subgroup_index(plist, r["subgroup"])
                return (idx, ) + tuple(start_sort_key(r))
            items.sort(key=sort_key)

            for r in items:
                if subgroup_index(plist, r["subgroup"]) >= 10_000:
                    msg = f"  [!] объект «{r['subgroup']}» (группа {GROUP_LABELS[g]}) не найден в справочнике приоритетов"
                    if msg not in unknown_warnings:
                        unknown_warnings.append(msg)

        ordered[g] = items

    if unknown_warnings:
        print("Предупреждения о неизвестных объектах:")
        for m in unknown_warnings:
            print(m)

    return ordered


# ------------------------------------------------------ СБОРКА ВЫХОДНОГО XLSX

def pick_month_year(records: list[dict], override_year: int | None) -> tuple[int, int]:
    """Определяет доминирующий месяц в заявках; год — по аргументу или по текущему."""
    months = Counter()
    years = Counter()
    for r in records:
        if r["start"]:
            y, m, _ = r["start"]
            months[m] += 1
            years[y] += 1
    month = months.most_common(1)[0][0] if months else datetime.now().month
    if override_year:
        year = override_year
    elif years:
        year = years.most_common(1)[0][0]
    else:
        year = datetime.now().year
    return month, year


def find_style_rows(ws_komi: Worksheet) -> dict:
    """Находит в проекте Коми РДУ подходящие строки-образцы для стилей."""
    header_last, data_last, sig_start = find_data_bounds(ws_komi)
    section_style_row = None
    equipment_style_row = None
    for r in range(header_last + 1, data_last + 1):
        if section_style_row is None and is_section_row(ws_komi, r):
            section_style_row = r
        if equipment_style_row is None and is_equipment_row(ws_komi, r):
            equipment_style_row = r
        if section_style_row and equipment_style_row:
            break
    return {
        "header_last": header_last,
        "data_last": data_last,
        "sig_start": sig_start,
        "section_style_row": section_style_row,
        "equipment_style_row": equipment_style_row,
    }


def write_header(ws_komi: Worksheet, out_ws: Worksheet, header_last: int):
    """Копирует шапку (строки 1..header_last) из Коми проекта в выходной лист."""
    for r in range(1, header_last + 1):
        copy_row_full(ws_komi, r, out_ws, r)
    # Объединения в пределах шапки.
    for mr in ws_komi.merged_cells.ranges:
        if mr.min_row <= header_last and mr.max_row <= header_last:
            rng = f"{get_column_letter(mr.min_col)}{mr.min_row}:{get_column_letter(min(mr.max_col, TABLE_COLS + 1))}{mr.max_row}"
            try:
                out_ws.merge_cells(rng)
            except Exception:
                pass


def write_title(out_ws: Worksheet, month: int, year: int):
    """Обновляет текст заголовка в шапке: «на <месяц> <год> г.»"""
    # Заголовок лежит в объединённой ячейке на 3-й строке (обычно C3:X3).
    # Поищем ячейку, значение которой начинается с 'Сводный график ремонта'.
    for r in range(1, 7):
        for c in range(1, TABLE_COLS + 1):
            v = out_ws.cell(r, c).value
            if isinstance(v, str) and v.strip().startswith("Сводный график"):
                new = (
                    "Сводный график ремонта ЛЭП и сетевого оборудования "
                    "операционных зон Архангельского и Коми РДУ "
                    f"на {RU_MONTHS_NOM[month]} {year} г."
                )
                out_ws.cell(r, c).value = new
                return


def write_style_row(out_ws: Worksheet, row: int, text: str,
                    src_ws: Worksheet, style_row: int,
                    height: float | None = None):
    """Пишет строку-заголовок/подзаголовок на всю ширину таблицы, копируя
    стиль из строки-образца проекта. Если указан `height` — принудительно
    выставляет высоту строки (pt); иначе копирует высоту из образца."""
    for c in range(1, TABLE_COLS + 1):
        copy_cell_style(src_ws.cell(style_row, c), out_ws.cell(row, c))
    out_ws.cell(row, 1).value = text
    rng = f"A{row}:{LAST_COL_LETTER}{row}"
    try:
        out_ws.merge_cells(rng)
    except Exception:
        pass
    if height is not None:
        out_ws.row_dimensions[row].height = height
    else:
        rh = src_ws.row_dimensions[style_row].height
        if rh is not None:
            out_ws.row_dimensions[row].height = rh


# ---------------------------------------------------------------------------
# Текстовая нормализация полей H (причины/условия) и N (вид ремонта, АГ, ...)
# ---------------------------------------------------------------------------
#
# Правила сформированы на основе сопоставления ручного сводного графика мая
# 2026 г. с исходными проектами. Каждое правило декларативно — регулярка +
# человекочитаемое имя (для отчёта) + действие.
#
# Действия:
#   * «H → N»  — короткая пометка в H вырезается и дописывается в конец N.
#   * «H drop» — короткая пометка в H вырезается без переноса.
#   * «ночь»   — «с включением на ночь» / «без включения на ночь» всегда
#                убирается из H и (если ещё нет) дописывается к N.
#   * «simple» — подстановки, применяемые и к H, и к N (общие нормализации).
#   * «преамбула» — опциональное сворачивание «Вывод в ремонт … для проведения
#                   <род. падеж> Y» → «<именительный падеж> Y» (флаг
#                   --collapse-preamble).

@dataclass
class NormOptions:
    """Настройки текстовой нормализации."""
    enabled: bool = True
    collapse_preamble: bool = False
    dry_run: bool = False


@dataclass
class NormStats:
    """Счётчики и детальный лог изменений для отчёта."""
    counts: Counter = field(default_factory=Counter)
    changes: list = field(default_factory=list)


# --- (1) Фразы из H, которые переносятся в конец N --------------------------

H_MOVE_TO_N_RULES: list[tuple[str, re.Pattern]] = [
    ("H→N «с переводом на ОШВ»", re.compile(r"с\s+переводом\s+на\s+ОШВ",
                                            re.IGNORECASE | re.UNICODE)),
    ("H→N «с переводом на ОВ»",  re.compile(r"с\s+переводом\s+на\s+ОВ",
                                            re.IGNORECASE | re.UNICODE)),
    ("H→N «Совместно с …»",       re.compile(r"Совместно\s+с\s+.+",
                                             re.IGNORECASE | re.UNICODE | re.DOTALL)),
]


# --- (2) Короткие «мусорные» ремарки, которые просто удаляются из H ---------

H_DROP_RULES: list[tuple[str, re.Pattern]] = [
    ("H убрано «не в транзите»",
     re.compile(r"не\s+в\s+транзите", re.IGNORECASE | re.UNICODE)),
    ("H убрано «с отключением без разбоки разъединителями»",
     re.compile(r"с\s+отключением\s+без\s+разб[оё]ки\s+разъединителями",
                re.IGNORECASE | re.UNICODE)),
]


# --- (3) Ночной режим — всегда переезжает из H в N --------------------------

NIGHT_RULES: list[tuple[str, re.Pattern]] = [
    ("«с включением на ночь» → N",
     re.compile(r"с\s+включением\s+на\s+ночь",  re.IGNORECASE | re.UNICODE)),
    ("«без включения на ночь» → N",
     re.compile(r"без\s+включения\s+на\s+ночь", re.IGNORECASE | re.UNICODE)),
]


# --- (4) Общие подстановки (применяются к H и N) ---------------------------

SIMPLE_SUBS: list[tuple[str, re.Pattern, object]] = [
    ("ТДТ → точки деления транзита",
     re.compile(r"\bТДТ\b", re.UNICODE), "точки деления транзита"),
    ("«А.Г.: ВЗ» → «А.Г.: ВЗ.»",
     re.compile(r"А\.Г\.:\s*ВЗ(?=\s)", re.UNICODE), "А.Г.: ВЗ."),
    ("«Включить» → «Включение»",
     re.compile(r"\bВключить\b", re.UNICODE), "Включение"),
    ("«Вывести в ремонт» → «Вывод в ремонт»",
     re.compile(r"\bВывести\s+в\s+ремонт\b", re.UNICODE), "Вывод в ремонт"),
    ("«NNNNг» → «NNNN г.»",
     re.compile(r"(\d{4})\s*г(?![а-яА-Я\.])", re.UNICODE), r"\1 г."),
    # Длинное тире между словами/числами (с пробелами вокруг).
    # Не трогает составные обозначения вида «АТ-2», «ВЛ-125», «28-30.05».
    ("Дефис между словами → длинное тире",
     re.compile(r"(?<=[A-Za-zА-Яа-я0-9])\s-\s(?=[A-Za-zА-Яа-я0-9])", re.UNICODE),
     " – "),
    # Пробел между числом и «кВ»: «110кВ» → «110 кВ».
    ("«NкВ» → «N кВ»",
     re.compile(r"(\d+)кВ\b", re.UNICODE), r"\1 кВ"),
    # Пробел между числом и «ч.»: «2ч» / «2ч.» → «2 ч.».
    # Не трогаем «часть», «часа», «часах» (следующая буква — русская).
    ("«Nч» → «N ч.»",
     re.compile(r"(?<![а-яА-Я\d])(\d+)\s*ч\.?(?![а-яА-Яa-zA-Z])", re.UNICODE),
     r"\1 ч."),
    # «ч. 30 м» / «ч.30 м» / «ч. 30м» → «ч. 30 мин.» — строго в контексте времени.
    ("«ч. Nм» → «ч. N мин.»",
     re.compile(r"(ч\.?)\s*(\d+)\s*м\.?(?![а-яА-Я])", re.UNICODE),
     r"\1 \2 мин."),
]


# --- (5) Опциональный коллапс преамбул в N ---------------------------------

PREAMBLE_RE = re.compile(
    r"(?P<prefix>.*?)"
    r"Вывод\w*\s+в\s+ремонт\s+"
    r"(?P<obj>.+?)"
    r"\s+(?P<link>на\s+время\s+проведения(?:\s+работ\s+по)?|"
    r"для\s+проведения|для|на\s+время)\s+"
    r"(?P<rest>.+)",
    re.IGNORECASE | re.UNICODE | re.DOTALL,
)

GEN_TO_NOM: dict[str, tuple[str, str]] = {
    "текущего ремонта":                 ("Текущий ремонт",       "текущий ремонт"),
    "среднего ремонта":                 ("Средний ремонт",       "средний ремонт"),
    "капитального ремонта":             ("Капитальный ремонт",   "капитальный ремонт"),
    "технического обслуживания":        ("Техническое обслуживание", "техническое обслуживание"),
    "профилактического восстановления": ("Профилактическое восстановление", "профилактическое восстановление"),
    "профилактическому восстановлению": ("Профилактическое восстановление", "профилактическое восстановление"),
    "испытаний":                        ("Проведение испытаний", "проведение испытаний"),
}


def _apply_h_rules(h: str, stats: NormStats) -> tuple[str, list[str]]:
    """Вычёркивает из H все предусмотренные фрагменты.
    Возвращает обновлённый H и список фрагментов для дописывания в N."""
    moves: list[str] = []
    text = h or ""

    # Ночь — всегда переносим
    for label, rx in NIGHT_RULES:
        m = rx.search(text)
        while m:
            moves.append(m.group(0))
            stats.counts[label] += 1
            text = text[:m.start()] + text[m.end():]
            m = rx.search(text)

    # Короткие хвосты H → N
    for label, rx in H_MOVE_TO_N_RULES:
        m = rx.search(text)
        if m:
            moves.append(m.group(0))
            stats.counts[label] += 1
            text = text[:m.start()] + text[m.end():]

    # Удаляемые пометки
    for label, rx in H_DROP_RULES:
        m = rx.search(text)
        if m:
            stats.counts[label] += 1
            text = text[:m.start()] + text[m.end():]

    # Чистка хвостов / повторов пробелов, но БЕЗ схлопывания пустых строк
    # между абзацами (пользователь мог поставить их осознанно).
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = text.strip()
    if re.fullmatch(r"[\s\.,;:]*", text or ""):
        text = ""

    return text, moves


def _append_moves_to_note(n: str, moves: list[str]) -> str:
    """Дописывает в конец N перенесённые из H фрагменты (без дублей)."""
    if not moves:
        return n or ""
    result = (n or "").rstrip()
    lowered = result.lower()
    for frag in moves:
        frag_norm = re.sub(r"\s+", " ", frag).strip()
        if not frag_norm:
            continue
        if frag_norm.lower() in lowered:
            continue
        if not result:
            result = frag_norm
        elif result.endswith(".") or result.endswith(":"):
            result = result + " " + frag_norm
        else:
            result = result + ". " + frag_norm
        lowered = result.lower()
    return result


def _apply_simple_subs(s: str, stats: NormStats) -> str:
    """Применяет список SIMPLE_SUBS к строке. Учитывает в статистике только
    фактические изменения (регекс может матчиться и на уже корректном тексте —
    такие «тождественные» срабатывания не считаем)."""
    if not s:
        return s
    for label, rx, repl in SIMPLE_SUBS:
        # Считаем, сколько матчей действительно меняют текст.
        if isinstance(repl, str):
            n_changed = sum(
                1 for m in rx.finditer(s)
                if m.expand(repl) != m.group(0)
            )
        else:
            n_changed = sum(
                1 for m in rx.finditer(s)
                if repl(m) != m.group(0)
            )
        if n_changed == 0:
            continue
        s = rx.sub(repl, s)
        stats.counts[label] += n_changed
    return s


def _collapse_preamble(n: str, stats: NormStats) -> str:
    """Сворачивает «Вывод в ремонт … для проведения <род> Y» в «<имен.> Y»."""
    if not n:
        return n
    m = PREAMBLE_RE.match(n)
    if not m:
        return n
    rest = m.group("rest")
    leading_key = None
    for k in sorted(GEN_TO_NOM.keys(), key=len, reverse=True):
        if re.match(r"^" + re.escape(k) + r"\b", rest, re.IGNORECASE):
            leading_key = k
            break
    if leading_key is None:
        return n

    stats.counts["Свёрнута преамбула «Вывод в ремонт … для проведения …»"] += 1
    cap_form, _ = GEN_TO_NOM[leading_key]
    rest_after = rest[len(leading_key):]
    # Остальные совпадения в rest_after приводим к lowercase-форме.
    for k in sorted(GEN_TO_NOM.keys(), key=len, reverse=True):
        lc_form = GEN_TO_NOM[k][1]
        rest_after = re.sub(r"\b" + re.escape(k) + r"\b", lc_form,
                            rest_after, flags=re.IGNORECASE)
    result = m.group("prefix") + cap_form + rest_after
    result = re.sub(r"[ \t]+", " ", result).rstrip()
    if not result.endswith("."):
        result += "."
    return result


def normalize_cells(h: str, n: str, opts: NormOptions, stats: NormStats,
                    row_label: str) -> tuple[str, str]:
    """Главная функция нормализации. Возвращает (new_H, new_N).

    Если opts.enabled = False — возвращает исходные значения без изменений.
    Все сработавшие правила учитываются в stats.counts; при любом изменении
    строка добавляется в stats.changes (для отчёта --dry-run)."""
    if not opts.enabled:
        return h or "", n or ""

    orig_h, orig_n = h or "", n or ""

    new_h, moves = _apply_h_rules(orig_h, stats)
    new_n = _append_moves_to_note(orig_n, moves)

    new_h = _apply_simple_subs(new_h, stats)
    new_n = _apply_simple_subs(new_n, stats)

    if opts.collapse_preamble:
        new_n = _collapse_preamble(new_n, stats)

    # Не логируем «пустое ≈ пустое» как изменение.
    def _changed(a: str, b: str) -> bool:
        if a == b:
            return False
        if (a or "").strip() == "" and (b or "").strip() == "":
            return False
        return True

    if _changed(orig_h, new_h) or _changed(orig_n, new_n):
        stats.changes.append({
            "row_label": row_label,
            "h_before":  orig_h,  "h_after": new_h,
            "n_before":  orig_n,  "n_after": new_n,
        })

    return new_h, new_n


def write_equipment_row(out_ws: Worksheet, dst_row: int, rec: dict,
                        opts: NormOptions, stats: NormStats):
    """Копирует строку оборудования из исходного листа, сохраняя стили и
    внутристрочные объединения (A:D для названия, N:O для примечания и т.п.),
    после чего нормализует текстовые поля H и N."""
    src_ws = rec["src_ws"]
    src_row = rec["src_row"]
    copy_row_full(src_ws, src_row, out_ws, dst_row)
    copy_merges_in_row(src_ws, src_row, out_ws, dst_row)

    h_cell = out_ws.cell(dst_row, 8)
    n_cell = out_ws.cell(dst_row, 14)
    row_label = f"R{dst_row} «{str(rec.get('name', '') or '')[:48]}»"
    new_h, new_n = normalize_cells(
        str(h_cell.value) if h_cell.value is not None else "",
        str(n_cell.value) if n_cell.value is not None else "",
        opts, stats, row_label,
    )
    if new_h != (h_cell.value or ""):
        h_cell.value = new_h if new_h else None
    if new_n != (n_cell.value or ""):
        n_cell.value = new_n if new_n else None

    # Гарантируем перенос текста в H и N (для авто-подгонки высоты строки).
    for cell in (h_cell, n_cell):
        al = cell.alignment
        if not al.wrap_text:
            cell.alignment = Alignment(
                horizontal=al.horizontal, vertical=al.vertical,
                text_rotation=al.text_rotation, wrap_text=True,
                shrink_to_fit=al.shrink_to_fit, indent=al.indent,
            )
    # Высоту строки данных не фиксируем — пусть Excel подгоняет сам.
    out_ws.row_dimensions[dst_row].height = None


def write_signatures(ws_komi: Worksheet, out_ws: Worksheet,
                     sig_start: int, dst_start: int) -> int:
    """Переносит блок подписей из Коми РДУ после итоговых строк данных.
    Возвращает индекс строки после последнего перенесённого ряда."""
    sig_end = ws_komi.max_row
    for i, r in enumerate(range(sig_start, sig_end + 1)):
        dst_r = dst_start + i
        copy_row_full(ws_komi, r, out_ws, dst_r)
        copy_merges_in_row(ws_komi, r, out_ws, dst_r)
    return dst_start + (sig_end - sig_start + 1)


def write_toc(out_ws: Worksheet, toc_row: int,
              group_anchors: dict[str, int]) -> None:
    """Пишет в строке `toc_row` оглавление: по ячейке на каждую непустую
    группу с гиперссылкой на строку её заголовка."""
    if not group_anchors:
        return

    ordered = [g for g in GROUP_ORDER if g in group_anchors]
    n = len(ordered)
    if n == 0:
        return

    # Равномерно распределяем непустые группы по 25 колонкам.
    base_width = TABLE_COLS // n
    extra = TABLE_COLS - base_width * n
    spans: list[tuple[int, int]] = []
    col = 1
    for i in range(n):
        w = base_width + (1 if i < extra else 0)
        spans.append((col, col + w - 1))
        col += w

    link_font = Font(name="Calibri", size=11, bold=True, color="0563C1",
                     underline="single")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    out_ws.row_dimensions[toc_row].height = ROW_HEIGHT_TOC
    for (lo, hi), g in zip(spans, ordered):
        anchor = group_anchors[g]
        cell = out_ws.cell(toc_row, lo)
        cell.value = f"{GROUP_LABELS[g]} (стр. {anchor})"
        cell.font = link_font
        cell.alignment = center
        cell.hyperlink = Hyperlink(
            ref=cell.coordinate,
            location=f"Page1!A{anchor}",
            display=cell.value,
        )
        if hi > lo:
            rng = (f"{get_column_letter(lo)}{toc_row}:"
                   f"{get_column_letter(hi)}{toc_row}")
            try:
                out_ws.merge_cells(rng)
            except Exception:
                pass


def _vid_remonta(n_text: str) -> str:
    """Извлекает короткий код вида ремонта из текста N (ВПр/ТР/СР/КР/ИСП/ЗРР/БВР).
    Возвращает пустую строку, если код не распознан."""
    if not n_text:
        return ""
    m = re.match(r"\s*(ВПр|ТР|СР|КР|ИСП|ЗРР|БВР)\b", n_text)
    return m.group(1) if m else ""


def _gantt_day_span(rec: dict, month: int, year: int,
                    scale_start: datetime, scale_end: datetime
                    ) -> tuple[int, int] | None:
    """По start/end записи возвращает (колонка_нач, колонка_кон) на шкале Ганта
    (индексы от 1) относительно scale_start. None — если дат нет или они
    полностью вне шкалы."""
    s, e = rec.get("start"), rec.get("end")
    if not s and not e:
        return None
    if s:
        sd = datetime(s[0], s[1], s[2])
    else:
        sd = datetime(year, month, 1)
    if e:
        ed = datetime(e[0], e[1], e[2])
    else:
        ed = sd
    if ed < sd:
        sd, ed = ed, sd
    if ed < scale_start or sd > scale_end:
        return None
    if sd < scale_start:
        sd = scale_start
    if ed > scale_end:
        ed = scale_end
    col_start = (sd - scale_start).days + 1
    col_end   = (ed - scale_start).days + 1
    return col_start, col_end


def build_gantt_sheet(out_wb: openpyxl.Workbook, gantt_items: list[dict],
                      month: int, year: int) -> None:
    """Добавляет в книгу лист «Диаграмма» с Гант-календарём.
    `gantt_items` — список словарей {row, group, rec}, в том же порядке, что
    записи на основном листе."""
    ws = out_wb.create_sheet(GANTT_SHEET_NAME)

    if not gantt_items:
        ws.cell(1, 1).value = "Нет строк для диаграммы."
        return

    # --- Шкала: от самой ранней start до самой поздней end, но гарантированно
    # включаем весь целевой месяц.
    month_start = datetime(year, month, 1)
    month_end = datetime(year, month, month_day_count(year, month))
    scale_start = month_start
    scale_end = month_end
    for it in gantt_items:
        s, e = it["rec"].get("start"), it["rec"].get("end")
        if s:
            d = datetime(s[0], s[1], s[2])
            if d < scale_start:
                scale_start = d
        if e:
            d = datetime(e[0], e[1], e[2])
            if d > scale_end:
                scale_end = d

    n_days = (scale_end - scale_start).days + 1

    COL_NAME = 1       # A — имя объекта
    COL_VID  = 2       # B — код вида ремонта
    COL_DAYS = 3       # C — первый день шкалы
    last_days_col = COL_DAYS + n_days - 1
    legend_col = last_days_col + 2   # пустая колонка-разрыв + легенда

    ws.column_dimensions[get_column_letter(COL_NAME)].width = 44
    ws.column_dimensions[get_column_letter(COL_VID)].width = 7
    for c in range(COL_DAYS, last_days_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 3.2

    thin_font = Font(name="Calibri", size=9)
    bold_font = Font(name="Calibri", size=10, bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    left   = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # --- Строка 1: месяцы (объединённые) ---
    cur = scale_start.replace(day=1)
    while cur <= scale_end:
        if cur.month == 12:
            nxt = cur.replace(year=cur.year + 1, month=1)
        else:
            nxt = cur.replace(month=cur.month + 1)
        seg_start = max(cur, scale_start)
        seg_end = min(nxt - timedelta(days=1), scale_end)
        col_from = COL_DAYS + (seg_start - scale_start).days
        col_to = COL_DAYS + (seg_end - scale_start).days
        cell = ws.cell(1, col_from)
        cell.value = f"{RU_MONTHS_SHORT[cur.month]} {cur.year}"
        cell.font = bold_font
        cell.alignment = center
        if col_to > col_from:
            rng = (f"{get_column_letter(col_from)}1:"
                   f"{get_column_letter(col_to)}1")
            try:
                ws.merge_cells(rng)
            except Exception:
                pass
        cur = nxt

    # --- Строка 2: числа дней + Строка 3: день недели ---
    weekend_fill = PatternFill(start_color=GANTT_COLOR_WEEKEND,
                               end_color=GANTT_COLOR_WEEKEND,
                               fill_type="solid")
    wday_names = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    for i in range(n_days):
        d = scale_start + timedelta(days=i)
        col = COL_DAYS + i
        c2 = ws.cell(2, col)
        c2.value = d.day
        c2.font = thin_font
        c2.alignment = center
        c3 = ws.cell(3, col)
        c3.value = wday_names[d.weekday()]
        c3.font = thin_font
        c3.alignment = center
        if d.weekday() >= 5:
            c2.fill = weekend_fill
            c3.fill = weekend_fill

    # Заголовки A/B.
    ws.cell(1, COL_NAME).value = "Объект"
    ws.cell(1, COL_NAME).font = bold_font
    ws.cell(1, COL_NAME).alignment = center
    ws.cell(1, COL_VID).value = "Вид"
    ws.cell(1, COL_VID).font = bold_font
    ws.cell(1, COL_VID).alignment = center
    try:
        ws.merge_cells(f"{get_column_letter(COL_NAME)}1:"
                       f"{get_column_letter(COL_NAME)}3")
        ws.merge_cells(f"{get_column_letter(COL_VID)}1:"
                       f"{get_column_letter(COL_VID)}3")
    except Exception:
        pass

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 14

    # --- Строки данных ---
    row = 4
    for it in gantt_items:
        rec = it["rec"]
        name = str(rec.get("name") or "").strip()
        sub = str(rec.get("subgroup") or "").strip()
        display = f"{sub}: {name}" if sub and sub.lower() != name.lower() else name

        n_text = ""
        try:
            n_text = str(rec["src_ws"].cell(rec["src_row"], 14).value or "")
        except Exception:
            pass
        vid = _vid_remonta(n_text)

        ws.cell(row, COL_NAME).value = display
        ws.cell(row, COL_NAME).font = thin_font
        ws.cell(row, COL_NAME).alignment = left
        ws.cell(row, COL_VID).value = vid or "—"
        ws.cell(row, COL_VID).font = thin_font
        ws.cell(row, COL_VID).alignment = center

        # Подкрашиваем выходные в строке данных тоже.
        for i in range(n_days):
            d = scale_start + timedelta(days=i)
            if d.weekday() >= 5:
                ws.cell(row, COL_DAYS + i).fill = weekend_fill

        span = _gantt_day_span(rec, month, year, scale_start, scale_end)
        if span:
            color = GANTT_COLORS.get(vid, GANTT_COLOR_OTHER)
            fill = PatternFill(start_color=color, end_color=color,
                               fill_type="solid")
            for c in range(COL_DAYS + span[0] - 1, COL_DAYS + span[1]):
                ws.cell(row, c).fill = fill
        row += 1

    # --- Легенда ---
    ws.cell(1, legend_col).value = "Легенда"
    ws.cell(1, legend_col).font = bold_font
    ws.cell(1, legend_col).alignment = center
    legend_rows = [
        ("ТР",  "Текущий ремонт"),
        ("СР",  "Средний ремонт"),
        ("КР",  "Капитальный ремонт"),
        ("ВПр", "Внеплановый ремонт"),
        ("ИСП", "Испытания"),
        ("ЗРР", "Заявка РР"),
        ("БВР", "Без вывода в ремонт"),
        ("—",   "Прочее / код не распознан"),
    ]
    for i, (code, desc) in enumerate(legend_rows):
        r = 2 + i
        color = GANTT_COLORS.get(code, GANTT_COLOR_OTHER)
        c1 = ws.cell(r, legend_col)
        c1.value = code
        c1.font = thin_font
        c1.alignment = center
        c1.fill = PatternFill(start_color=color, end_color=color,
                              fill_type="solid")
        c2 = ws.cell(r, legend_col + 1)
        c2.value = desc
        c2.font = thin_font
        c2.alignment = left
    ws.column_dimensions[get_column_letter(legend_col)].width = 6
    ws.column_dimensions[get_column_letter(legend_col + 1)].width = 28

    # Закрепление областей: под шапкой и справа от столбцов-идентификаторов.
    ws.freeze_panes = ws.cell(4, COL_DAYS).coordinate

    # Печать — альбомная, вписать в 1 страницу по ширине.
    try:
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    ws.print_area = (f"A1:{get_column_letter(legend_col + 1)}{row - 1}")


def build_output(priority: dict, records: list[dict],
                 ws_komi: Worksheet, ws_arkh: Worksheet | None,
                 month: int, year: int,
                 opts: NormOptions, stats: NormStats,
                 apply_sort: bool = True,
                 apply_toc: bool = True,
                 apply_heights: bool = True,
                 apply_gantt: bool = True) -> openpyxl.Workbook:
    """Собирает итоговую книгу.

    Флаги-стадии позволяют выключить отдельные преобразования:
      * apply_sort    — группировать и сортировать по приоритетам. Если False —
                        строки идут в порядке `records`, но классификация (для
                        заголовков групп) выполняется всегда.
      * apply_toc     — писать строку-оглавление со ссылками на группы.
      * apply_heights — фиксировать высоту строк-заголовков (секций/подсекций).
      * apply_gantt   — создавать лист «Диаграмма».
    Нормализация текста управляется `opts.enabled`.
    """
    style_info = find_style_rows(ws_komi)

    out_wb = openpyxl.Workbook()
    # удаляем стандартный лист «Sheet» и создаём «Page1», чтобы совпадало с проектами
    out_wb.remove(out_wb.active)
    out_ws = out_wb.create_sheet("Page1")

    # Ширины колонок — как в проекте Коми РДУ.
    copy_column_widths(ws_komi, out_ws, ncols=26)
    # параметры страницы
    try:
        out_ws.page_setup = _copy(ws_komi.page_setup)
        out_ws.print_options = _copy(ws_komi.print_options)
        out_ws.page_margins = _copy(ws_komi.page_margins)
        out_ws.sheet_properties.pageSetUpPr = _copy(ws_komi.sheet_properties.pageSetUpPr)
    except Exception:
        pass

    # Шапка.
    write_header(ws_komi, out_ws, style_info["header_last"])
    write_title(out_ws, month, year)

    # Группированные записи.
    if apply_sort:
        grouped = group_and_sort(records, priority)
    else:
        # Только классификация, без переупорядочивания: сохраняем порядок
        # строк из `records`, но раскладываем по корзинам согласно classify().
        buckets: "OrderedDict[str, list[dict]]" = OrderedDict()
        for rec in records:
            g, sub = classify(rec)
            rec["group"] = g
            rec["subgroup"] = sub
            buckets.setdefault(g, []).append(rec)
        grouped = OrderedDict()
        for g in GROUP_ORDER:
            if g in buckets and buckets[g]:
                grouped[g] = buckets[g]

    sect_style_row = style_info["section_style_row"]
    # Резервируем строку под оглавление; фактический текст TOC запишем в конце,
    # когда будут известны позиции всех заголовков групп.
    toc_row = style_info["header_last"] + 1
    cur = toc_row + 1

    group_anchors: dict[str, int] = {}
    # Порядок записей в том же виде, в котором они идут на листе (для Гант-листа).
    gantt_items: list[dict] = []

    section_h = ROW_HEIGHT_SECTION if apply_heights else None
    subsection_h = ROW_HEIGHT_SUBSECTION if apply_heights else None

    for g in GROUP_ORDER:
        if g not in grouped or not grouped[g]:
            continue
        # Заголовок группы.
        group_anchors[g] = cur
        write_style_row(out_ws, cur, GROUP_LABELS[g], ws_komi, sect_style_row,
                        height=section_h)
        cur += 1

        items = grouped[g]

        if g in (GROUP_PS220, GROUP_PS110, GROUP_ES, GROUP_OGR):
            # Подзаголовки по объектам.
            current_sub = None
            for r in items:
                if r["subgroup"] != current_sub:
                    current_sub = r["subgroup"]
                    if current_sub:
                        write_style_row(out_ws, cur, current_sub, ws_komi,
                                        sect_style_row,
                                        height=subsection_h)
                        cur += 1
                write_equipment_row(out_ws, cur, r, opts, stats)
                gantt_items.append({"row": cur, "group": g, "rec": r})
                cur += 1
        else:
            # «Плоские» группы (ЛЭП, АЧР, Прочее).
            for r in items:
                write_equipment_row(out_ws, cur, r, opts, stats)
                gantt_items.append({"row": cur, "group": g, "rec": r})
                cur += 1

    # Оглавление (гиперссылки на строки заголовков групп).
    if apply_toc:
        write_toc(out_ws, toc_row, group_anchors)

    # Подписи.
    write_signatures(ws_komi, out_ws, style_info["sig_start"], cur)

    # установим область печати (A..Y)
    out_ws.print_area = f"A1:{LAST_COL_LETTER}{out_ws.max_row}"

    # Второй лист — Гант-календарь.
    if apply_gantt:
        build_gantt_sheet(out_wb, gantt_items, month, year)

    # Основной лист должен открываться первым.
    out_wb.active = 0

    return out_wb


# ---------------------------------------------------------------------------
# Работа с уже существующим сводником: парсер, резервные копии, inplace-стадии
# ---------------------------------------------------------------------------


def find_existing_svod(root: Path) -> Path | None:
    """Возвращает путь к последнему (по mtime) файлу «Сводный график …xlsx»
    в корне проекта. None — если файла нет."""
    candidates = [
        p for p in root.glob(f"{SVOD_FILE_PREFIX}*.xlsx")
        if not p.name.startswith("~$")  # временный lock-файл Excel
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def make_backup(path: Path, log=print) -> Path | None:
    """Кладёт копию `path` в `backups/<timestamp>__<имя>.xlsx`. Возвращает путь
    к копии. Если файла нет — возвращает None."""
    if not path.exists():
        return None
    BACKUP_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dst = BACKUP_DIR / f"{ts}__{path.name}"
    shutil.copy2(path, dst)
    log(f"Резервная копия: backups/{dst.name}")
    return dst


def restore_latest_backup(svod_path: Path, log=print) -> Path | None:
    """Восстанавливает `svod_path` из последней подходящей копии в `backups/`.
    Возвращает путь к восстановленному файлу либо None, если копий нет."""
    if not BACKUP_DIR.exists():
        log("Папка backups/ не найдена — нечего восстанавливать.")
        return None
    prefix = svod_path.name
    candidates = sorted(
        BACKUP_DIR.glob(f"*__{prefix}"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        log(f"В backups/ нет копий «{prefix}».")
        return None
    latest = candidates[0]
    # Сначала — копия текущего файла (на случай неудачного отката).
    if svod_path.exists():
        make_backup(svod_path, log=log)
    shutil.copy2(latest, svod_path)
    log(f"Восстановлено из: backups/{latest.name}")
    return svod_path


def is_toc_row(ws: Worksheet, row: int) -> bool:
    """Строка-оглавление в сводника: хотя бы одна ячейка содержит гиперссылку
    с локацией «Page1!A…»."""
    for c in range(1, TABLE_COLS + 1):
        cell = ws.cell(row, c)
        hl = getattr(cell, "hyperlink", None)
        if hl is None:
            continue
        loc = getattr(hl, "location", None) or ""
        if "Page1" in str(loc):
            return True
    return False


def extract_records_from_svod(ws_svod: Worksheet, default_year: int,
                              src_key: str = "svod") -> list[dict]:
    """Аналог `extract_records`, но для уже сгенерированного сводника.

    Пропускает строку оглавления и строки-заголовки групп (таких подгрупп
    нет в проектах, они появляются только в своднике). Подзаголовки по
    объектам (ПС/Электростанции) становятся current_section — как и в
    исходных проектах."""
    header_last, data_last, _sig_start = find_data_bounds(ws_svod)
    recs: list[dict] = []
    current_section = ""
    group_label_lc = {v.strip().lower() for v in GROUP_LABELS.values()}

    for r in range(header_last + 1, data_last + 1):
        if is_toc_row(ws_svod, r):
            continue
        a = ws_svod.cell(r, 1).value
        if a is None or (isinstance(a, str) and a.strip() == ""):
            continue
        name = str(a).strip()

        if is_section_row(ws_svod, r):
            if name.strip().lower() in group_label_lc:
                # Заголовок группы верхнего уровня — данных не содержит.
                current_section = ""
                continue
            current_section = name
            continue

        # Определяем, из какого РДУ запись. В своднике это признак косвенный:
        # секция-подзаголовок может содержать «Архангельского» / «Коми».
        rdu = "Коми"
        sec_lc = current_section.lower()
        if "арх" in sec_lc:
            rdu = "Арх"
        elif "коми" in sec_lc:
            rdu = "Коми"

        start_raw = ws_svod.cell(r, 6).value
        end_raw = ws_svod.cell(r, 7).value
        start = parse_day_month(start_raw, default_year)
        end = parse_day_month(end_raw, default_year)

        recs.append({
            "rdu": rdu,
            "section": current_section,
            "name": name,
            "start": start,
            "end": end,
            "src_ws": ws_svod,
            "src_row": r,
            "src_key": src_key,
        })
    return recs


def _save_with_backup(wb: openpyxl.Workbook, out_path: Path, log=print):
    """Сохраняет книгу по пути `out_path`, предварительно забэкапив старый
    файл (если он был). Понятно отрабатывает PermissionError."""
    if out_path.exists():
        make_backup(out_path, log=log)
    try:
        wb.save(out_path)
    except PermissionError:
        raise RuntimeError(
            f"Не удаётся сохранить «{out_path.name}» — вероятно, файл открыт "
            f"в Excel. Закройте его и попробуйте ещё раз."
        )
    log(f"Сохранено: {out_path.name}")


# --- Стадии, выполняемые «поверх» уже существующего сводника -----------------

def stage_normalize_inplace(svod_path: Path, opts: NormOptions,
                            stats: NormStats, log=print) -> None:
    """Нормализует текст в столбцах H/N непосредственно в файле `svod_path`."""
    log(f"Нормализация текста: {svod_path.name}")
    wb = openpyxl.load_workbook(svod_path)
    if "Page1" not in wb.sheetnames:
        raise RuntimeError("В файле нет листа «Page1» — это не похоже на сводный график.")
    ws = wb["Page1"]
    header_last, data_last, _ = find_data_bounds(ws)
    opts.enabled = True  # стадия явно включает нормализацию
    for r in range(header_last + 1, data_last + 1):
        if is_toc_row(ws, r) or is_section_row(ws, r):
            continue
        if not is_equipment_row(ws, r):
            continue
        h_cell = ws.cell(r, 8)
        n_cell = ws.cell(r, 14)
        row_label = f"R{r} «{_short(str(ws.cell(r, 1).value or ''), 48)}»"
        new_h, new_n = normalize_cells(
            str(h_cell.value) if h_cell.value is not None else "",
            str(n_cell.value) if n_cell.value is not None else "",
            opts, stats, row_label,
        )
        if new_h != (h_cell.value or ""):
            h_cell.value = new_h if new_h else None
        if new_n != (n_cell.value or ""):
            n_cell.value = new_n if new_n else None
    _save_with_backup(wb, svod_path, log=log)


def stage_build_toc_inplace(svod_path: Path, log=print) -> None:
    """Пересоздаёт строку оглавления в уже существующем своднике."""
    log(f"Оглавление: {svod_path.name}")
    wb = openpyxl.load_workbook(svod_path)
    if "Page1" not in wb.sheetnames:
        raise RuntimeError("В файле нет листа «Page1».")
    ws = wb["Page1"]
    header_last, data_last, _ = find_data_bounds(ws)

    # Найдём якоря: строки-секции, текст которых совпадает с названием одной
    # из групп верхнего уровня.
    group_anchors: dict[str, int] = {}
    label_to_key = {v.strip().lower(): k for k, v in GROUP_LABELS.items()}
    for r in range(header_last + 1, data_last + 1):
        if not is_section_row(ws, r):
            continue
        text = str(ws.cell(r, 1).value or "").strip().lower()
        key = label_to_key.get(text)
        if key and key not in group_anchors:
            group_anchors[key] = r

    toc_row = header_last + 1
    # Сначала — unmerge в строке TOC (иначе ячейки в merged-диапазоне — read-only).
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row == toc_row and mr.max_row == toc_row:
            ws.unmerge_cells(str(mr))
    # Затем чистим значения и гиперссылки старого оглавления.
    for c in range(1, TABLE_COLS + 1):
        cell = ws.cell(toc_row, c)
        cell.value = None
        cell.hyperlink = None

    write_toc(ws, toc_row, group_anchors)
    _save_with_backup(wb, svod_path, log=log)


def stage_set_heights_inplace(svod_path: Path, log=print) -> None:
    """Фиксирует высоты строк-заголовков в уже существующем своднике и
    включает wrap_text в столбцах H/N."""
    log(f"Фиксация высот + wrap: {svod_path.name}")
    wb = openpyxl.load_workbook(svod_path)
    if "Page1" not in wb.sheetnames:
        raise RuntimeError("В файле нет листа «Page1».")
    ws = wb["Page1"]
    header_last, data_last, _ = find_data_bounds(ws)
    label_set = {v.strip().lower() for v in GROUP_LABELS.values()}

    toc_row = header_last + 1
    if is_toc_row(ws, toc_row):
        ws.row_dimensions[toc_row].height = ROW_HEIGHT_TOC

    for r in range(header_last + 1, data_last + 1):
        if not is_section_row(ws, r):
            continue
        text = str(ws.cell(r, 1).value or "").strip().lower()
        if text in label_set:
            ws.row_dimensions[r].height = ROW_HEIGHT_SECTION
        else:
            ws.row_dimensions[r].height = ROW_HEIGHT_SUBSECTION

    # wrap_text для H/N в строках данных + авто-подгонка высоты.
    for r in range(header_last + 1, data_last + 1):
        if is_toc_row(ws, r) or is_section_row(ws, r):
            continue
        if not is_equipment_row(ws, r):
            continue
        for col in (8, 14):
            cell = ws.cell(r, col)
            al = cell.alignment
            if not al.wrap_text:
                cell.alignment = Alignment(
                    horizontal=al.horizontal, vertical=al.vertical,
                    text_rotation=al.text_rotation, wrap_text=True,
                    shrink_to_fit=al.shrink_to_fit, indent=al.indent,
                )
        ws.row_dimensions[r].height = None

    _save_with_backup(wb, svod_path, log=log)


def stage_build_gantt_inplace(svod_path: Path, default_year: int,
                              log=print) -> None:
    """Пересоздаёт лист «Диаграмма» в уже существующем своднике."""
    log(f"Диаграмма Ганта: {svod_path.name}")
    wb = openpyxl.load_workbook(svod_path)
    if "Page1" not in wb.sheetnames:
        raise RuntimeError("В файле нет листа «Page1».")
    ws = wb["Page1"]

    recs = extract_records_from_svod(ws, default_year=default_year)
    # Классифицируем, чтобы gantt_items имели subgroup (для имени строки).
    for rec in recs:
        g, sub = classify(rec)
        rec["group"] = g
        rec["subgroup"] = sub or rec.get("section", "")
    month, year = pick_month_year(recs, default_year if default_year else None)

    if GANTT_SHEET_NAME in wb.sheetnames:
        del wb[GANTT_SHEET_NAME]

    gantt_items = [{"row": r["src_row"], "group": r["group"], "rec": r}
                   for r in recs]
    build_gantt_sheet(wb, gantt_items, month, year)
    wb.active = wb.index(wb["Page1"])
    _save_with_backup(wb, svod_path, log=log)


# --- «Большие» стадии: полная пересборка ------------------------------------

def _load_inputs(root: Path, year_hint: int | None, log=print
                 ) -> tuple[dict, list[dict], Worksheet, int, int]:
    """Загружает справочник и проекты, возвращает (priority, records,
    template_ws, month, year). Падает с RuntimeError при проблемах."""
    p_prio = find_file(FILE_PRIO)
    p_arkh = find_file(FILE_ARKH)
    p_komi = find_file(FILE_KOMI)

    if not p_prio:
        raise RuntimeError(
            f"Не найден файл справочника «{FILE_PRIO}».\n"
            f"Положите его в папку: {root}"
        )
    if not p_arkh and not p_komi:
        raise RuntimeError(
            f"Не найдены ни «{FILE_ARKH}», ни «{FILE_KOMI}».\n"
            f"Положите хотя бы один из них в папку: {root}"
        )

    log("Найдены файлы:")
    log(f"  • {p_prio}")
    if p_arkh:
        log(f"  • {p_arkh}")
    if p_komi:
        log(f"  • {p_komi}")

    priority = load_priority(p_prio)
    default_year = year_hint if year_hint else datetime.now().year
    records: list[dict] = []
    ws_arkh = ws_komi = None

    if p_arkh:
        wb_arkh = openpyxl.load_workbook(p_arkh)
        ws_arkh = wb_arkh["Page1"] if "Page1" in wb_arkh.sheetnames else wb_arkh.active
        validate_project_template(ws_arkh, p_arkh.name)
        records += extract_records(ws_arkh, "Арх", default_year, "arkh")

    if p_komi:
        wb_komi = openpyxl.load_workbook(p_komi)
        ws_komi = wb_komi["Page1"] if "Page1" in wb_komi.sheetnames else wb_komi.active
        validate_project_template(ws_komi, p_komi.name)
        records += extract_records(ws_komi, "Коми", default_year, "komi")

    log(f"Всего строк оборудования: {len(records)}")
    template_ws = ws_komi or ws_arkh
    month, year = pick_month_year(records, year_hint)
    log(f"Месяц сводного: {RU_MONTHS_NOM[month]} {year}")
    return priority, records, template_ws, month, year


def stage_full_rebuild(root: Path, year_hint: int | None,
                       opts: NormOptions, stats: NormStats,
                       log=print,
                       apply_sort: bool = True,
                       apply_toc: bool = True,
                       apply_heights: bool = True,
                       apply_gantt: bool = True) -> Path:
    """Собирает сводный график «с нуля» из проектов, со всеми выбранными
    стадиями. Возвращает путь к сохранённому файлу."""
    priority, records, template_ws, month, year = _load_inputs(
        root, year_hint, log=log)
    out_wb = build_output(
        priority, records, template_ws, None, month, year, opts, stats,
        apply_sort=apply_sort, apply_toc=apply_toc,
        apply_heights=apply_heights, apply_gantt=apply_gantt,
    )
    out_name = (
        f"{SVOD_FILE_PREFIX} ЛЭП и сетевого оборудования "
        f"на {RU_MONTHS_NOM[month]} {year} г.xlsx"
    )
    out_path = root / out_name
    _save_with_backup(out_wb, out_path, log=log)
    return out_path


def stage_rebuild_from_existing(svod_path: Path, year_hint: int | None,
                                opts: NormOptions, stats: NormStats,
                                log=print) -> Path:
    """Перечитывает существующий сводник и перестраивает его (полный набор
    стадий: расстановка приоритетов + TOC + высоты + Гант + нормализация,
    управляемая `opts`). Справочник приоритетов нужен обязательно.

    Стили и подписи берутся из самого сводника — он же и шаблон."""
    p_prio = find_file(FILE_PRIO)
    if not p_prio:
        raise RuntimeError(
            f"Не найден файл справочника «{FILE_PRIO}».\n"
            f"Положите его в папку: {ROOT}"
        )
    priority = load_priority(p_prio)

    log(f"Читаем существующий сводник: {svod_path.name}")
    wb = openpyxl.load_workbook(svod_path)
    if "Page1" not in wb.sheetnames:
        raise RuntimeError("В файле нет листа «Page1».")
    ws = wb["Page1"]

    default_year = year_hint if year_hint else datetime.now().year
    records = extract_records_from_svod(ws, default_year=default_year)
    log(f"Строк оборудования в своднике: {len(records)}")

    month, year = pick_month_year(records, year_hint)

    out_wb = build_output(
        priority, records, ws, None, month, year, opts, stats,
        apply_sort=True, apply_toc=True, apply_heights=True, apply_gantt=True,
    )
    # Имя файла оставляем тем же (месяц/год могут чуть измениться — тогда
    # возьмём новое имя). Бэкап старого выполнится в _save_with_backup.
    out_name = (
        f"{SVOD_FILE_PREFIX} ЛЭП и сетевого оборудования "
        f"на {RU_MONTHS_NOM[month]} {year} г.xlsx"
    )
    out_path = svod_path.parent / out_name
    # Если имя совпадает со старым — перезаписываем; если нет — старый тоже
    # бэкапим, чтобы не плодить разные копии.
    if out_path != svod_path and svod_path.exists():
        make_backup(svod_path, log=log)
    _save_with_backup(out_wb, out_path, log=log)
    return out_path


# ---------------------------------------------------------------- ТОЧКА ВХОДА

def _short(s: str, limit: int = 100) -> str:
    """Укорачивает многострочный текст до одной строки ≤ limit символов."""
    if s is None:
        return ""
    s = str(s).replace("\n", " ⏎ ")
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > limit:
        s = s[: limit - 1] + "…"
    return s


def _print_norm_report(stats: NormStats, dry_run: bool) -> None:
    """Печатает отчёт о применённых правилах нормализации текста."""
    print()
    print("Нормализация текста:")
    if not stats.counts:
        print("  изменений нет.")
    else:
        for label, c in sorted(stats.counts.items(), key=lambda kv: (-kv[1], kv[0])):
            print(f"  • {label}: {c}")

    if dry_run:
        print()
        print("Детализация изменений (--dry-run, файл не сохранён):")
        if not stats.changes:
            print("  нет.")
        for ch in stats.changes:
            print(f"  {ch['row_label']}")
            if ch['h_before'] != ch['h_after']:
                print(f"    H: {_short(ch['h_before'])}")
                print(f"     →  {_short(ch['h_after']) or '(пусто)'}")
            if ch['n_before'] != ch['n_after']:
                print(f"    N: {_short(ch['n_before'])}")
                print(f"     →  {_short(ch['n_after']) or '(пусто)'}")


STAGE_CHOICES = (
    "all",         # полная пересборка с нуля (по умолчанию)
    "merge",       # только объединение проектов (без сортировки/TOC/Ганта/высот)
    "sort",        # перечитать существующий сводник и переставить по приоритетам
    "normalize",   # только нормализация текста в готовом своднике
    "toc",         # только перегенерация оглавления
    "heights",     # только фиксация высот и wrap_text
    "gantt",       # только перестроить лист «Диаграмма»
    "restore",     # откатить сводник к последней резервной копии
)


def _require_existing_svod(log=print) -> Path:
    """Возвращает путь к существующему своднику в корне или падает с понятной
    ошибкой."""
    path = find_existing_svod(ROOT)
    if path is None:
        raise RuntimeError(
            f"В папке {ROOT} не найден файл «{SVOD_FILE_PREFIX} …xlsx».\n"
            f"Сначала выполните стадию «merge» или «all»."
        )
    return path


def main():
    parser = argparse.ArgumentParser(
        description="Сборщик сводного графика ремонтов ЛЭП и сетевого оборудования."
    )
    parser.add_argument("--stage", choices=STAGE_CHOICES, default="all",
                        help="Какую стадию выполнить. По умолчанию «all» — "
                             "полная пересборка из проектов.")
    parser.add_argument("--year", type=int, default=None,
                        help="Год в имени выходного файла (по умолчанию — из дат проекта или текущий).")
    parser.add_argument("--no-normalize", action="store_true",
                        help="Отключить текстовую нормализацию полей H и N "
                             "(применяется к «all», «merge», «sort»).")
    parser.add_argument("--collapse-preamble", action="store_true",
                        help="Сворачивать преамбулы «Вывод в ремонт … для проведения …» "
                             "в краткую форму «<Вид ремонта> Y» (опытное правило).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Ничего не сохранять — только показать, что будет изменено "
                             "(работает для «all» и «merge»).")
    args = parser.parse_args()

    opts = NormOptions(
        enabled=not args.no_normalize,
        collapse_preamble=bool(args.collapse_preamble),
        dry_run=bool(args.dry_run),
    )
    stats = NormStats()

    try:
        if args.stage == "all":
            if opts.dry_run:
                # «Сухой прогон» собираем в памяти, но не сохраняем.
                priority, records, tws, month, year = _load_inputs(
                    ROOT, args.year, log=print)
                build_output(priority, records, tws, None, month, year,
                             opts, stats)
                _print_norm_report(stats, dry_run=True)
                print()
                print("[--dry-run] Итоговый файл не сохранён.")
                return
            out_path = stage_full_rebuild(
                ROOT, args.year, opts, stats,
                apply_sort=True, apply_toc=True,
                apply_heights=True, apply_gantt=True,
            )
            _print_norm_report(stats, dry_run=False)
            print(f"\nГотово: {out_path}")

        elif args.stage == "merge":
            # Чистое объединение: классификация без приоритетов, без TOC/высот/Ганта.
            out_path = stage_full_rebuild(
                ROOT, args.year, opts, stats,
                apply_sort=False, apply_toc=False,
                apply_heights=False, apply_gantt=False,
            )
            _print_norm_report(stats, dry_run=False)
            print(f"\nГотово: {out_path}")

        elif args.stage == "sort":
            svod = _require_existing_svod()
            out_path = stage_rebuild_from_existing(
                svod, args.year, opts, stats, log=print)
            _print_norm_report(stats, dry_run=False)
            print(f"\nГотово: {out_path}")

        elif args.stage == "normalize":
            svod = _require_existing_svod()
            stage_normalize_inplace(svod, opts, stats, log=print)
            _print_norm_report(stats, dry_run=False)

        elif args.stage == "toc":
            svod = _require_existing_svod()
            stage_build_toc_inplace(svod, log=print)

        elif args.stage == "heights":
            svod = _require_existing_svod()
            stage_set_heights_inplace(svod, log=print)

        elif args.stage == "gantt":
            svod = _require_existing_svod()
            stage_build_gantt_inplace(svod,
                                      args.year or datetime.now().year,
                                      log=print)

        elif args.stage == "restore":
            svod = find_existing_svod(ROOT)
            if svod is None:
                # Нечего откатывать в корне — попробуем восстановить по любой
                # копии: возьмём ту, к чьему имени больше всего копий.
                raise RuntimeError(
                    f"В папке {ROOT} не найден свод. Сначала положите файл "
                    f"«{SVOD_FILE_PREFIX} …xlsx» или запустите стадию «merge»/"
                    f"«all»."
                )
            restored = restore_latest_backup(svod, log=print)
            if restored is None:
                sys.exit(4)

    except RuntimeError as e:
        print(f"ОШИБКА: {e}")
        sys.exit(2)


if __name__ == "__main__":
    main()

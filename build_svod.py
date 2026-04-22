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
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
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
                    src_ws: Worksheet, style_row: int):
    """Пишет строку-заголовок/подзаголовок на всю ширину таблицы, копируя
    стиль из строки-образца проекта."""
    for c in range(1, TABLE_COLS + 1):
        copy_cell_style(src_ws.cell(style_row, c), out_ws.cell(row, c))
    out_ws.cell(row, 1).value = text
    rng = f"A{row}:{LAST_COL_LETTER}{row}"
    try:
        out_ws.merge_cells(rng)
    except Exception:
        pass
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
    if not s:
        return s
    for label, rx, repl in SIMPLE_SUBS:
        new, n_subs = rx.subn(repl, s)
        if n_subs:
            stats.counts[label] += n_subs
            s = new
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


def build_output(priority: dict, records: list[dict],
                 ws_komi: Worksheet, ws_arkh: Worksheet | None,
                 month: int, year: int,
                 opts: NormOptions, stats: NormStats) -> openpyxl.Workbook:
    """Собирает итоговую книгу."""
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
    grouped = group_and_sort(records, priority)

    sect_style_row = style_info["section_style_row"]
    cur = style_info["header_last"] + 1

    for g in GROUP_ORDER:
        if g not in grouped or not grouped[g]:
            continue
        # Заголовок группы.
        write_style_row(out_ws, cur, GROUP_LABELS[g], ws_komi, sect_style_row)
        cur += 1

        items = grouped[g]

        if g in (GROUP_PS220, GROUP_PS110, GROUP_ES, GROUP_OGR):
            # Подзаголовки по объектам.
            current_sub = None
            for r in items:
                if r["subgroup"] != current_sub:
                    current_sub = r["subgroup"]
                    if current_sub:
                        write_style_row(out_ws, cur, current_sub, ws_komi, sect_style_row)
                        cur += 1
                write_equipment_row(out_ws, cur, r, opts, stats)
                cur += 1
        else:
            # «Плоские» группы (ЛЭП, АЧР, Прочее).
            for r in items:
                write_equipment_row(out_ws, cur, r, opts, stats)
                cur += 1

    # Подписи.
    write_signatures(ws_komi, out_ws, style_info["sig_start"], cur)

    # установим область печати (A..Y)
    out_ws.print_area = f"A1:{LAST_COL_LETTER}{out_ws.max_row}"

    return out_wb


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


def main():
    parser = argparse.ArgumentParser(
        description="Сборщик сводного графика ремонтов ЛЭП и сетевого оборудования."
    )
    parser.add_argument("--year", type=int, default=None,
                        help="Год в имени выходного файла (по умолчанию — из дат проекта или текущий).")
    parser.add_argument("--no-normalize", action="store_true",
                        help="Отключить текстовую нормализацию полей H и N.")
    parser.add_argument("--collapse-preamble", action="store_true",
                        help="Сворачивать преамбулы «Вывод в ремонт … для проведения …» "
                             "в краткую форму «<Вид ремонта> Y» (опытное правило).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Ничего не сохранять — только показать, что будет изменено.")
    args = parser.parse_args()

    opts = NormOptions(
        enabled=not args.no_normalize,
        collapse_preamble=bool(args.collapse_preamble),
        dry_run=bool(args.dry_run),
    )
    stats = NormStats()

    p_prio = find_file(FILE_PRIO)
    p_arkh = find_file(FILE_ARKH)
    p_komi = find_file(FILE_KOMI)

    if not p_prio:
        print(f"ОШИБКА: не найден файл справочника «{FILE_PRIO}».")
        print(f"Положите его в папку: {ROOT}")
        sys.exit(1)

    if not p_arkh and not p_komi:
        print(f"ОШИБКА: не найдены ни «{FILE_ARKH}», ни «{FILE_KOMI}».")
        print(f"Положите хотя бы один из них в папку: {ROOT}")
        sys.exit(1)

    print("Найдены файлы:")
    print(f"  • {p_prio}")
    if p_arkh: print(f"  • {p_arkh}")
    if p_komi: print(f"  • {p_komi}")

    priority = load_priority(p_prio)

    default_year = args.year if args.year else datetime.now().year
    records: list[dict] = []
    ws_arkh = ws_komi = None

    if p_arkh:
        wb_arkh = openpyxl.load_workbook(p_arkh)
        ws_arkh = wb_arkh["Page1"] if "Page1" in wb_arkh.sheetnames else wb_arkh.active
        records += extract_records(ws_arkh, "Арх", default_year, "arkh")

    if p_komi:
        wb_komi = openpyxl.load_workbook(p_komi)
        ws_komi = wb_komi["Page1"] if "Page1" in wb_komi.sheetnames else wb_komi.active
        records += extract_records(ws_komi, "Коми", default_year, "komi")

    print(f"Всего строк оборудования: {len(records)}")

    # если Коми нет — используем шаблон из Арх
    template_ws = ws_komi or ws_arkh

    month, year = pick_month_year(records, args.year)
    print(f"Месяц сводного: {RU_MONTHS_NOM[month]} {year}")

    out_wb = build_output(priority, records, template_ws, ws_arkh, month, year,
                          opts, stats)

    out_name = (
        f"Сводный график ремонтов ЛЭП и сетевого оборудования "
        f"на {RU_MONTHS_NOM[month]} {year} г.xlsx"
    )
    out_path = ROOT / out_name

    _print_norm_report(stats, dry_run=opts.dry_run)

    if opts.dry_run:
        print()
        print("[--dry-run] Итоговый файл не сохранён.")
        return

    try:
        out_wb.save(out_path)
    except PermissionError:
        print(f"ОШИБКА: не удаётся сохранить «{out_path.name}» — вероятно, файл открыт в Excel.")
        print("Закройте его и запустите скрипт ещё раз.")
        sys.exit(2)

    print()
    print(f"Готово: {out_path}")


if __name__ == "__main__":
    main()

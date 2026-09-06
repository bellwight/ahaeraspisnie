#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер расписания НГОК для одной группы.

Работает напрямую с XML внутри xlsx (без тяжёлой построчной загрузки
через openpyxl в read_only-режиме, которая на файле с сотней листов
и объединёнными ячейками работает очень медленно на случайном доступе).

Логика:
1. Скачиваем xlsx по прямой ссылке (Google Sheets export).
2. Во всех листах ищем ячейки, где встречается точное название нужной
   группы (заголовок блока колонок "Дисциплина/Преподаватель/Ауд."
   для этой группы на конкретной неделе).
3. Для каждого найденного блока читаем даты (колонка A) и время
   (колонка B) с учётом объединённых ячеек и вытаскиваем пары.
4. Все найденные пары (по всем листам/неделям) объединяем, убираем
   дубли, сортируем по дате и времени.
"""
import re
import json
import zipfile
import shutil
import datetime
import urllib.request
from pathlib import Path

from openpyxl.utils import range_boundaries, column_index_from_string
from openpyxl.utils.datetime import from_excel

GROUP_NAME = "01.26.Р.ИИ.ГД.ОФ.9-НСК"

STANDARD_SLOTS = [
    "9:00-10:30",
    "10:40-12:10",
    "12:30-14:00",
    "14:20-15:50",
    "16:00-17:30",
    "17:40-19:10",
]

TIME_RE = re.compile(r"^\d{1,2}[.\-]\d{2}-\d{1,2}\.\d{2}$")
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(?:(.*?)</c>|)|<c r="([A-Z]+)(\d+)"([^>]*)/>', re.S)


def is_time_text(v):
    return isinstance(v, str) and bool(TIME_RE.match(v.strip()))


def is_subject_text(v):
    return isinstance(v, str) and "(" in v and v.strip() not in ("", "№")


def is_name_text(v):
    if not isinstance(v, str):
        return False
    v = v.strip()
    if not v or "(" in v:
        return False
    return len(v.split()) >= 2


def download_xlsx(url: str, dest: Path):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as resp, open(dest, "wb") as f:
        shutil.copyfileobj(resp, f)


def load_shared_strings(z: zipfile.ZipFile):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    shared = z.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
    si_blocks = re.findall(r"<si>(.*?)</si>", shared, re.S)
    strings = []
    for block in si_blocks:
        text = "".join(re.findall(r"<t[^>]*>(.*?)</t>", block, re.S))
        text = re.sub(r"<[^>]+>", "", text)
        text = (text.replace("&amp;", "&").replace("&lt;", "<")
                .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
        strings.append(text)
    return strings


def find_group_cells(z: zipfile.ZipFile, shared_strings, group_name: str):
    target_indices = {i for i, s in enumerate(shared_strings) if s.strip() == group_name}
    results = []
    if not target_indices:
        return results

    wb_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")
    rel_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(worksheets/sheet\d+\.xml)"', rels_xml))
    sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"[^>]*/>', wb_xml)

    for sheet_name, rid in sheets:
        target = rel_map.get(rid)
        if not target or ("xl/" + target) not in z.namelist():
            continue
        content = z.read("xl/" + target).decode("utf-8", errors="ignore")
        for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*t="s"[^>]*><v>(\d+)</v></c>', content):
            col_letter, row_str, val = m.groups()
            if int(val) in target_indices:
                results.append({
                    "sheet_name": sheet_name,
                    "sheet_file": target,
                    "row": int(row_str),
                    "col": column_index_from_string(col_letter),
                })
    return results


def get_merges(content: str):
    merges_raw = re.findall(r'<mergeCell ref="([^"]+)"/>', content)
    merge_map = {}
    for m in merges_raw:
        min_col, min_row, max_col, max_row = range_boundaries(m)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                merge_map[(c, r)] = (min_col, min_row)
    return merge_map


_INNER_T = re.compile(r"<t[^>]*>(.*?)</t>", re.S)
_INNER_V = re.compile(r"<v>(.*?)</v>", re.S)


def _cell_value(attrs: str, inner: str, shared_strings):
    t_match = re.search(r't="([^"]+)"', attrs)
    ctype = t_match.group(1) if t_match else None
    if inner is None:
        return None
    if ctype == "s":
        vm = _INNER_V.search(inner)
        if not vm:
            return None
        idx = int(vm.group(1))
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else None
    if ctype in ("str", "inlineStr"):
        tm = _INNER_T.search(inner)
        if tm:
            txt = tm.group(1)
            txt = (txt.replace("&amp;", "&").replace("&lt;", "<")
                   .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
            return txt
        return None
    # numeric or date (style-based); return raw float, caller decides
    vm = _INNER_V.search(inner)
    if not vm:
        return None
    raw = vm.group(1)
    try:
        return float(raw)
    except ValueError:
        return raw


def read_window(content: str, shared_strings, row_start: int, row_end: int, cols_wanted):
    """
    Возвращает dict {(col,row): value} для строк [row_start, row_end]
    и колонок из cols_wanted. Значения дат остаются числами (float) -
    конвертация в datetime делается отдельно только для колонки дат,
    т.к. без разбора styles.xml не всегда легко однозначно отличить
    дату от обычного числа, а нам известно, что даты встречаются
    только в колонке 1.
    """
    result = {}
    # ограничим область поиска по регулярке диапазоном строк с помощью
    # быстрого среза текста: ищем строки "<row r="N" ...>" в диапазоне
    row_iter = re.finditer(r'<row[^>]*\br="(\d+)"[^>]*>(.*?)</row>', content, re.S)
    for rm in row_iter:
        rnum = int(rm.group(1))
        if rnum < row_start:
            continue
        if rnum > row_end:
            break
        row_body = rm.group(2)
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', row_body, re.S):
            col_letter, _r, attrs, inner = cm.groups()
            col = column_index_from_string(col_letter)
            if col not in cols_wanted:
                continue
            val = _cell_value(attrs, inner, shared_strings)
            result[(col, rnum)] = val
    return result


def extract_block(content, shared_strings, merge_map, subj_col, header_row, max_scan_rows=400):
    room_col = subj_col + 3
    cols_wanted = {1, 2, subj_col, room_col}
    raw = read_window(content, shared_strings, header_row, header_row + max_scan_rows, cols_wanted)

    cache = {}

    def getval(r, c):
        key = (c, r)
        tl = merge_map.get(key, key)
        if tl in cache:
            return cache[tl]
        v = raw.get(tl)
        cache[tl] = v
        return v

    def get_date(r, c):
        v = getval(r, c)
        if isinstance(v, float):
            try:
                return from_excel(v)
            except Exception:
                return None
        return None

    lessons = []
    row = header_row + 1
    end_row = header_row + max_scan_rows

    # Поиск преподавателя делаем не по границам сегмента/дня, а просто
    # в небольшой окрестности строки с названием предмета в ГЛОБАЛЬНОМ
    # окне (raw), т.к. из-за переноса строк с длинным текстом реальная
    # строка с ФИО преподавателя иногда попадает в соседний сегмент
    # или даже в блок следующего дня.
    def find_teacher(subj_row):
        for offset in (1, 2, -1, 3, -2):
            r = subj_row + offset
            v = raw.get((subj_col, r))
            if is_name_text(v):
                return v.strip()
        return None

    def make_lesson(date_val, slot_index, subj_row, room):
        subject = raw.get((subj_col, subj_row))
        if isinstance(room, str):
            room = room.strip()
        elif isinstance(room, float):
            room = str(int(room)) if room.is_integer() else str(room)
        return {
            "date": date_val.strftime("%Y-%m-%d"),
            "time": STANDARD_SLOTS[slot_index],
            "subject": subject.strip(),
            "teacher": find_teacher(subj_row),
            "room": room,
        }

    def flush_day(date_val, rows):
        """
        В этих таблицах порядковый номер пары ("№") не всегда стоит на
        одной строке со своим предметом: иногда он идёт ПЕРЕД
        предметом (обычный случай), иногда ПОСЛЕ (метка на самом деле
        завершает предыдущую пару, а не открывает новую), а иногда
        сам предмет вообще написан на одной строке с номером. При
        этом абсолютное значение номера "плывёт" ото дня ко дню
        (может начинаться не с 1). Поэтому:
        1) сперва одним проходом собираем список РАЗЛИЧНЫХ номеров
           пар в порядке первого появления - это и есть их порядковое
           место в дне (0 = первая пара 9:00-10:30 и т.д.);
        2) вторым проходом сопоставляем каждый найденный предмет с
           "своим" номером: тем же номером на этой строке, иначе
           ближайшим ещё не занятым номером из уже встреченных,
           иначе - следующим по ходу вперёд.
        """
        if date_val is None or not rows:
            return

        def label_at(r):
            v = raw.get((2, r))
            return v if isinstance(v, float) and v.is_integer() else None

        label_sequence = []
        for r in rows:
            lv = label_at(r)
            if lv is not None and lv not in label_sequence:
                label_sequence.append(lv)
        canonical_index = {lv: i for i, lv in enumerate(label_sequence)}

        from collections import deque
        unassigned = deque()
        consumed = set()

        def peek_next_label(after_row):
            for r2 in rows:
                if r2 <= after_row:
                    continue
                lv2 = label_at(r2)
                if lv2 is not None and lv2 not in consumed:
                    return lv2
            return None

        for r in rows:
            lv = label_at(r)
            v = getval(r, subj_col)
            has_subject = is_subject_text(v)

            if has_subject:
                if lv is not None:
                    chosen = lv
                    if chosen in unassigned:
                        unassigned.remove(chosen)
                elif unassigned:
                    chosen = unassigned.pop()
                else:
                    chosen = peek_next_label(r)
                if chosen is None or chosen not in canonical_index:
                    continue
                consumed.add(chosen)
                idx = canonical_index[chosen]
                if idx >= len(STANDARD_SLOTS):
                    continue
                room = getval(r, room_col)
                lessons.append(make_lesson(date_val, idx, r, room))
            elif lv is not None and lv not in consumed and lv not in unassigned:
                unassigned.append(lv)

    def infer_missing_leading_date(next_date, n_missing):
        """Первый день недели в этих таблицах иногда идёт без даты в
        колонке A. Восстанавливаем её отсчётом назад от следующей
        известной даты, пропуская воскресенья (в этом расписании
        пар по воскресеньям не бывает)."""
        d = next_date
        out = []
        for _ in range(n_missing):
            d = d - datetime.timedelta(days=1)
            while d.weekday() == 6:  # Sunday
                d = d - datetime.timedelta(days=1)
            out.append(d)
        out.reverse()
        return out

    current_date = None
    day_rows = []
    pending_undated_blocks = []  # список списков строк без даты, до первой найденной даты

    while row <= end_row:
        c2_raw = raw.get((2, row))
        if isinstance(c2_raw, str) and c2_raw.strip() == "№":
            break
        d_raw = raw.get((1, row))
        d = None
        if isinstance(d_raw, float):
            try:
                d = from_excel(d_raw)
            except Exception:
                d = None
        if d is not None:
            if current_date is None and day_rows:
                # это была первая (безымянная) неделя без даты
                pending_undated_blocks.append(day_rows)
            else:
                flush_day(current_date, day_rows)
            current_date = d
            day_rows = [row]
        else:
            day_rows.append(row)
        row += 1

    flush_day(current_date, day_rows)

    # обрабатываем безымянные дни в начале блока (обычно только понедельник)
    if pending_undated_blocks and current_date is not None:
        # используем самую первую известную дату всего блока как опору
        first_known_date = None
        for l in lessons:
            fd = datetime.date.fromisoformat(l["date"])
            if first_known_date is None or fd < first_known_date:
                first_known_date = fd
        if first_known_date is None:
            first_known_date = current_date
        inferred_dates = infer_missing_leading_date(first_known_date, len(pending_undated_blocks))
        for inferred_date, rows in zip(inferred_dates, pending_undated_blocks):
            flush_day(inferred_date, rows)

    lessons.sort(key=lambda l: (l["date"], l["time"]))
    return lessons


def parse_all(xlsx_path: Path, group_name: str = GROUP_NAME):
    all_lessons = []
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = load_shared_strings(z)
        cells = find_group_cells(z, shared_strings, group_name)

        content_cache = {}
        merge_cache = {}
        for c in cells:
            sheet_file = c["sheet_file"]
            path = "xl/" + sheet_file
            if sheet_file not in content_cache:
                content_cache[sheet_file] = z.read(path).decode("utf-8", errors="ignore")
                merge_cache[sheet_file] = get_merges(content_cache[sheet_file])
            content = content_cache[sheet_file]
            merge_map = merge_cache[sheet_file]
            lessons = extract_block(content, shared_strings, merge_map, c["col"], c["row"])
            all_lessons.extend(lessons)

    seen = set()
    dedup = []
    for l in all_lessons:
        key = (l["date"], l["time"], l["subject"])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(l)

    dedup.sort(key=lambda l: (l["date"], l["time"]))
    return dedup


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True, help="Прямая ссылка на скачивание xlsx")
    ap.add_argument("--group", default=GROUP_NAME)
    ap.add_argument("--out", default="docs/schedule.json")
    ap.add_argument("--keep-days", type=int, default=14,
                     help="сколько дней в прошлом хранить в выдаче")
    ap.add_argument("--local-file", default=None,
                     help="локальный xlsx вместо скачивания (для теста)")
    args = ap.parse_args()

    if args.local_file:
        tmp_xlsx = Path(args.local_file)
    else:
        tmp_xlsx = Path("/tmp/_source.xlsx")
        download_xlsx(args.url, tmp_xlsx)

    lessons = parse_all(tmp_xlsx, args.group)

    cutoff = (datetime.date.today() - datetime.timedelta(days=args.keep_days)).isoformat()
    lessons = [l for l in lessons if l["date"] >= cutoff]

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "group": args.group,
        "generated_at": datetime.datetime.utcnow().isoformat() + "Z",
        "lessons": lessons,
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: {len(lessons)} занятий -> {out_path}")


if __name__ == "__main__":
    main()

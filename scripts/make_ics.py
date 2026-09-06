#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерирует schedule.ics из schedule.json (по одному VEVENT на пару)."""
import json
import hashlib
import datetime
from pathlib import Path

TZID = "Asia/Novosibirsk"


def esc(text: str) -> str:
    if text is None:
        return ""
    return (text.replace("\\", "\\\\").replace(";", "\\;")
            .replace(",", "\\,").replace("\n", "\\n"))


def fold(line: str) -> str:
    """RFC5545: строки длиннее 75 октетов сворачиваются продолжением,
    начинающимся с пробела. Режем по байтам UTF-8, не разрывая
    многобайтовые последовательности посередине."""
    data = line.encode("utf-8")
    if len(data) <= 75:
        return line
    out = []
    limit = 75
    while len(data) > limit:
        cut = limit
        # не резать многобайтовый символ пополам (продолжающие байты 0b10xxxxxx)
        while cut > 0 and (data[cut] & 0xC0) == 0x80:
            cut -= 1
        out.append(data[:cut].decode("utf-8"))
        data = b" " + data[cut:]
        limit = 75
    out.append(data.decode("utf-8"))
    return "\r\n".join(out)


def build_ics(payload: dict) -> str:
    group = payload["group"]
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//ngok-schedule//" + esc(group) + "//RU",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:" + esc(f"Расписание {group}"),
        "X-WR-TIMEZONE:" + TZID,
        "BEGIN:VTIMEZONE",
        "TZID:" + TZID,
        "BEGIN:STANDARD",
        "DTSTART:19700101T000000",
        "TZOFFSETFROM:+0700",
        "TZOFFSETTO:+0700",
        "TZNAME:+07",
        "END:STANDARD",
        "END:VTIMEZONE",
    ]

    for lesson in payload["lessons"]:
        date_str = lesson["date"]
        time_str = lesson["time"]  # "9:00-10:30"
        start_s, end_s = time_str.split("-")

        def parse_hm(s):
            s = s.strip().replace(".", ":")
            h, m = s.split(":")
            return int(h), int(m)

        sh, sm = parse_hm(start_s)
        eh, em = parse_hm(end_s)
        y, mo, d = (int(x) for x in date_str.split("-"))
        dtstart = datetime.datetime(y, mo, d, sh, sm)
        dtend = datetime.datetime(y, mo, d, eh, em)

        uid_src = f"{group}|{date_str}|{time_str}|{lesson['subject']}"
        uid = hashlib.sha1(uid_src.encode("utf-8")).hexdigest() + "@ngok-schedule"

        summary = lesson["subject"]
        location = lesson.get("room") or ""
        desc_parts = []
        if lesson.get("teacher"):
            desc_parts.append(lesson["teacher"])
        desc_parts.append(group)
        description = "\\n".join(esc(p) for p in desc_parts)

        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{datetime.datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}")
        lines.append(f"DTSTART;TZID={TZID}:{dtstart.strftime('%Y%m%dT%H%M%S')}")
        lines.append(f"DTEND;TZID={TZID}:{dtend.strftime('%Y%m%dT%H%M%S')}")
        lines.append(fold("SUMMARY:" + esc(summary)))
        if location:
            lines.append(fold("LOCATION:" + esc(location)))
        lines.append(fold("DESCRIPTION:" + description))
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", default="docs/schedule.json")
    ap.add_argument("--out", default="docs/schedule.ics")
    args = ap.parse_args()

    payload = json.loads(Path(args.inp).read_text(encoding="utf-8"))
    ics = build_ics(payload)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(ics, encoding="utf-8")
    print(f"OK -> {out_path}")


if __name__ == "__main__":
    main()

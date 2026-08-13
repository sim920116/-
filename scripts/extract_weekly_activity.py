#!/usr/bin/env python3
"""주간보고 엑셀 파일에서 사육 활동(입식/폐사/출하)을 추출해 data/weekly_activity_<이름>.csv에 반영한다.

사용법:
    python scripts/extract_weekly_activity.py <이름> <엑셀파일경로> <weekOf>

    예) python scripts/extract_weekly_activity.py 심유선 ~/Downloads/주간보고_0811.xlsx 2026-08-11

동일한 weekOf로 재실행하면 기존 주차 데이터를 새 결과로 교체한다(중복 누적 방지).
"""

import argparse
import csv
import re
import sys
from datetime import datetime, date
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent

CSV_FIELDS = [
    "weekOf", "date", "group", "activity", "count",
    "weight_kg", "note", "source_file", "extracted_at",
]

# 헤더 텍스트에서 컬럼 성격을 판별하기 위한 키워드
DATE_KEYWORDS = ["날짜", "일자", "발생일"]
GROUP_KEYWORDS = ["그룹", "배치"]
WEIGHT_KEYWORDS = ["중량", "체중"]
NOTE_KEYWORDS = ["비고", "메모"]
KIND_KEYWORDS = ["구분", "활동", "유형"]  # 롱포맷: 이 컬럼 값이 활동명(폐사/출하/입식 등)
COUNT_KEYWORDS = ["두수", "수량"]  # 롱포맷에서 구분과 짝을 이루는 수량 컬럼

# 와이드포맷: 컬럼명에 아래 키워드가 있으면 그 컬럼 자체가 해당 활동의 두수
ACTIVITY_COLUMN_KEYWORDS = {
    "폐사": ["폐사"],
    "출하": ["출하", "판매"],
    "입식": ["입식", "전입"],
}


def normalize_header(cell_value) -> str:
    return str(cell_value).strip() if cell_value is not None else ""


def find_header_row(rows, max_scan=10):
    """앞부분 max_scan행 중 인식 가능한 키워드가 가장 많이 매칭되는 행을 헤더로 판단한다."""
    all_keywords = (
        DATE_KEYWORDS + GROUP_KEYWORDS + WEIGHT_KEYWORDS + NOTE_KEYWORDS
        + KIND_KEYWORDS + COUNT_KEYWORDS
        + [kw for kws in ACTIVITY_COLUMN_KEYWORDS.values() for kw in kws]
    )
    best_idx, best_score = None, 0
    for idx, row in enumerate(rows[:max_scan]):
        headers = [normalize_header(c) for c in row]
        score = sum(1 for h in headers if any(kw in h for kw in all_keywords))
        if score > best_score:
            best_idx, best_score = idx, score
    return best_idx


def classify_columns(headers):
    cols = {
        "date": None, "group": None, "weight": None, "note": None,
        "kind": None, "count": None, "activity": {},
    }
    for i, h in enumerate(headers):
        if not h:
            continue
        matched_activity = None
        for activity, kws in ACTIVITY_COLUMN_KEYWORDS.items():
            if any(kw in h for kw in kws):
                matched_activity = activity
                break

        if matched_activity is not None:
            # "폐사두수"처럼 두수/수량이 포함돼도 활동 전용 컬럼을 우선한다
            cols["activity"][matched_activity] = i
        elif cols["date"] is None and any(kw in h for kw in DATE_KEYWORDS):
            cols["date"] = i
        elif cols["group"] is None and any(kw in h for kw in GROUP_KEYWORDS):
            cols["group"] = i
        elif cols["weight"] is None and any(kw in h for kw in WEIGHT_KEYWORDS):
            cols["weight"] = i
        elif cols["note"] is None and any(kw in h for kw in NOTE_KEYWORDS):
            cols["note"] = i
        elif cols["kind"] is None and any(kw in h for kw in KIND_KEYWORDS):
            cols["kind"] = i
        elif cols["count"] is None and any(kw in h for kw in COUNT_KEYWORDS):
            cols["count"] = i
    return cols


def parse_date(value, fallback: str) -> str:
    if value is None or value == "":
        return fallback
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    m = re.match(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", text)
    if m:
        y, mo, d = m.groups()
        return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
    return fallback


def to_number(value):
    if value is None or value == "":
        return None
    try:
        n = float(value)
        return int(n) if n == int(n) else n
    except (ValueError, TypeError):
        return None


def extract_records(filepath: Path, week_of: str):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    records = []

    for sheet in wb.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue
        header_idx = find_header_row(rows)
        if header_idx is None:
            continue
        headers = [normalize_header(c) for c in rows[header_idx]]
        cols = classify_columns(headers)

        for row in rows[header_idx + 1:]:
            if row is None or all(c is None or str(c).strip() == "" for c in row):
                continue

            row_date = parse_date(row[cols["date"]] if cols["date"] is not None else None, week_of)
            group = normalize_header(row[cols["group"]]) if cols["group"] is not None else ""
            weight = to_number(row[cols["weight"]]) if cols["weight"] is not None else None
            note = normalize_header(row[cols["note"]]) if cols["note"] is not None else ""

            row_records = []

            # 롱포맷: 구분(활동명) + 두수 컬럼
            if cols["kind"] is not None and cols["count"] is not None:
                kind = normalize_header(row[cols["kind"]])
                count = to_number(row[cols["count"]])
                if kind and count:
                    row_records.append((kind, count))

            # 와이드포맷: 활동별 전용 컬럼
            for activity, idx in cols["activity"].items():
                count = to_number(row[idx])
                if count:
                    row_records.append((activity, count))

            for activity, count in row_records:
                records.append({
                    "weekOf": week_of,
                    "date": row_date,
                    "group": group,
                    "activity": activity,
                    "count": count,
                    "weight_kg": weight if weight is not None else "",
                    "note": note,
                    "source_file": filepath.name,
                    "extracted_at": datetime.now().isoformat(timespec="seconds"),
                })

    return records


def load_existing(csv_path: Path):
    if not csv_path.exists():
        return []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_csv(csv_path: Path, rows):
    def sort_key(r):
        return (r.get("weekOf", ""), r.get("date", ""), r.get("group", ""), r.get("activity", ""))

    rows = sorted(rows, key=sort_key)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in CSV_FIELDS})


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("name", help="사육자/농장 이름 (예: 심유선)")
    parser.add_argument("filepath", help="업로드된 주간보고 엑셀 파일 경로")
    parser.add_argument("weekOf", help="보고 대상 주차 시작일 (예: 2026-08-11)")
    args = parser.parse_args()

    filepath = Path(args.filepath).expanduser()
    if not filepath.exists():
        sys.exit(f"엑셀 파일을 찾을 수 없습니다: {filepath}")

    new_records = extract_records(filepath, args.weekOf)
    if not new_records:
        sys.exit(
            "추출된 활동 내역이 없습니다. 엑셀 헤더에 '그룹/배치', '폐사', '출하/판매', "
            "'입식/전입' 등의 컬럼명이 있는지 확인해주세요."
        )

    csv_path = REPO_ROOT / "data" / f"weekly_activity_{args.name}.csv"
    existing = load_existing(csv_path)
    # 같은 weekOf의 기존 기록은 새 결과로 교체(재실행 시 중복 누적 방지)
    kept = [r for r in existing if r.get("weekOf") != args.weekOf]
    write_csv(csv_path, kept + new_records)

    summary = {}
    for r in new_records:
        summary[r["activity"]] = summary.get(r["activity"], 0) + r["count"]
    summary_text = ", ".join(f"{k} {v}건" for k, v in summary.items())
    print(f"{args.weekOf} 주차 반영 완료 -> {csv_path.relative_to(REPO_ROOT)}")
    print(f"  {summary_text} (총 {len(new_records)}행)")


if __name__ == "__main__":
    main()

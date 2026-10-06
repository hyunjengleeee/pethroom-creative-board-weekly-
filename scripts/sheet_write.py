"""분석 결과를 시트에 쓰기 (같은 키가 있으면 덮어쓰고, 없으면 아래에 추가)

    python scripts/sheet_write.py tags   tags.json     # [{"광고 이름": ..., "포맷": ..., ...}]
    python scripts/sheet_write.py report report.json   # {"주 시작": "2026-09-28", "한 줄 요약": ..., ...}
"""
import json
import sys

from build_week import REPORT_COLS, TAG_COLS, open_sheet, tab

sys.stdout.reconfigure(encoding="utf-8")


def upsert(ws_title, header, rows):
    sh = open_sheet()
    tab(sh, ws_title, header)  # 없으면 만들기
    ws = sh.worksheet(ws_title)
    values = ws.get_all_values()
    head = values[0] if values else header
    where = {r[0].strip(): i + 1 for i, r in enumerate(values) if r and r[0].strip()}
    updates, appends = [], []
    for row in rows:
        line = [str(row.get(h, "")) for h in head]
        key = line[0].strip()
        if key in where:
            updates.append({"range": f"A{where[key]}", "values": [line]})
        else:
            appends.append(line)
    if updates:
        ws.batch_update(updates, value_input_option="RAW")
    if appends:
        ws.append_rows(appends, value_input_option="RAW")
    print(f"{ws_title}: 수정 {len(updates)} · 추가 {len(appends)}")


def main():
    kind, path = sys.argv[1], sys.argv[2]
    data = json.load(open(path, encoding="utf-8"))
    if kind == "tags":
        upsert("소재태그", TAG_COLS, data)
    elif kind == "report":
        upsert("주간리포트", REPORT_COLS, data if isinstance(data, list) else [data])
    else:
        raise SystemExit("tags 또는 report")


if __name__ == "__main__":
    main()

"""
Đẩy dữ liệu top holders lên Google Sheets:
- Mỗi token có 1 sheet "<TOKEN>_Log" dạng dài: Ngày | Địa chỉ ví | Số dư | Chênh lệch
  (mỗi lần chạy thêm N dòng mới, 1 dòng/ví — dễ lọc, pivot, vẽ biểu đồ)
- 1 sheet "Dashboard" tóm tắt: tổng cung top 100, số ví mới/rớt hạng,
  top ví tăng/giảm mạnh nhất, + 1 bảng lịch sử nhỏ (Ngày | Tổng cung mỗi token)
  để cắm biểu đồ xu hướng (chart set up 1 lần, tự động nhận dữ liệu mới vì
  tham chiếu nguyên cột).

Cần biến môi trường:
    GOOGLE_SHEETS_CREDENTIALS  - nội dung JSON của service account (dán nguyên văn)
    GOOGLE_SHEETS_ID           - ID của Google Sheet (lấy từ URL)
    TONAPI_KEY                 - (tùy chọn) API key TonAPI

Chạy thủ công:
    python push_to_gsheets.py
"""

import os
import json
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import (
    CellFormat, TextFormat, Color, format_cell_range, set_frozen,
    ConditionalFormatRule, BooleanRule, BooleanCondition, GridRange,
    get_conditional_format_rules,
)

from ton_data import TOKENS, fetch_token_snapshot

VN_TZ = timezone(timedelta(hours=7))
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADER_BG = Color(0.12, 0.31, 0.47)  # xanh đậm, đồng bộ với bản Excel
HEADER_FMT = CellFormat(
    backgroundColor=HEADER_BG,
    textFormat=TextFormat(bold=True, foregroundColor=Color(1, 1, 1)),
    horizontalAlignment="CENTER",
)


def get_client():
    creds_raw = os.environ["GOOGLE_SHEETS_CREDENTIALS"]
    info = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_or_create_worksheet(sh, title, rows=200, cols=10):
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def ensure_log_header(ws):
    values = ws.get_all_values()
    if not values or values[0][:4] != ["Ngày", "Địa chỉ ví", "Số dư", "Chênh lệch"]:
        ws.update("A1:D1", [["Ngày", "Địa chỉ ví", "Số dư", "Chênh lệch"]])
        format_cell_range(ws, "A1:D1", HEADER_FMT)
        set_frozen(ws, rows=1)
        # tô màu xanh/đỏ tự động cho cột Chênh lệch, áp dụng luôn cho các dòng
        # sẽ được thêm sau này (range mở rộng tới dòng 20000)
        rules = get_conditional_format_rules(ws)
        rules.clear()
        rules.append(ConditionalFormatRule(
            ranges=[GridRange(sheetId=ws.id, startRowIndex=1, endRowIndex=20000,
                               startColumnIndex=3, endColumnIndex=4)],
            booleanRule=BooleanRule(
                condition=BooleanCondition("NUMBER_GREATER", ["0"]),
                format=CellFormat(backgroundColor=Color(0.78, 0.94, 0.81)),
            ),
        ))
        rules.append(ConditionalFormatRule(
            ranges=[GridRange(sheetId=ws.id, startRowIndex=1, endRowIndex=20000,
                               startColumnIndex=3, endColumnIndex=4)],
            booleanRule=BooleanRule(
                condition=BooleanCondition("NUMBER_LESS", ["0"]),
                format=CellFormat(backgroundColor=Color(0.98, 0.78, 0.81)),
            ),
        ))
        rules.save()


def get_prev_balances(ws):
    """Đọc dòng của ngày gần nhất trong Log sheet -> {địa chỉ: số dư}."""
    values = ws.get_all_values()
    if len(values) <= 1:
        return {}, None
    last_date = values[-1][0]
    prev = {}
    for row in values[1:]:
        if row[0] == last_date and len(row) >= 3:
            try:
                prev[row[1]] = float(row[2])
            except ValueError:
                pass
    return prev, last_date


def append_log_rows(ws, date_label, holders):
    """holders: list[(address, balance)]. Trả về (rows_appended, new_wallets, dropped_wallets, movers)."""
    prev_balances, last_date = get_prev_balances(ws)
    today_balances = dict(holders)

    all_addrs = dict(today_balances)
    for addr in prev_balances:
        if addr not in all_addrs:
            all_addrs[addr] = 0

    rows = []
    movers = []  # (address, diff)
    new_wallets = 0
    dropped_wallets = 0
    for addr, bal in all_addrs.items():
        prev_val = prev_balances.get(addr)
        diff = (bal - prev_val) if prev_val is not None else ""
        rows.append([date_label, addr, round(bal, 4), round(diff, 4) if diff != "" else ""])
        if prev_val is None:
            new_wallets += 1
        elif bal == 0 and prev_val > 0:
            dropped_wallets += 1
        if diff != "":
            movers.append((addr, diff))

    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")

    movers.sort(key=lambda x: x[1], reverse=True)
    top_gainers = movers[:3]
    top_losers = sorted(movers, key=lambda x: x[1])[:3]

    total_today = sum(today_balances.values())
    total_prev = sum(prev_balances.values()) if prev_balances else None

    return {
        "new_wallets": new_wallets,
        "dropped_wallets": dropped_wallets,
        "top_gainers": top_gainers,
        "top_losers": top_losers,
        "total_today": total_today,
        "total_prev": total_prev,
        "wallet_count": len(today_balances),
    }


def update_dashboard(sh, date_label, summaries):
    """summaries: {token_name: summary_dict}"""
    ws = get_or_create_worksheet(sh, "Dashboard", rows=200, cols=12)

    title = f"TỔNG QUAN VÍ TOP 100 — cập nhật {date_label}"
    ws.update("A1", [[title]])
    format_cell_range(ws, "A1:F1", CellFormat(textFormat=TextFormat(bold=True, fontSize=14)))

    row = 3
    for token_name, s in summaries.items():
        block = [
            [f"── {token_name} ──", "", "", ""],
            ["Tổng cung top 100 hôm nay", round(s["total_today"], 2), "", ""],
            ["Chênh lệch so lần trước",
             round(s["total_today"] - s["total_prev"], 2) if s["total_prev"] is not None else "N/A", "", ""],
            ["Số ví đang theo dõi", s["wallet_count"], "", ""],
            ["Ví mới xuất hiện", s["new_wallets"], "", ""],
            ["Ví rớt khỏi top 100", s["dropped_wallets"], "", ""],
            ["Top 3 tăng mạnh nhất:", "", "", ""],
        ]
        for addr, diff in s["top_gainers"]:
            block.append(["", addr, round(diff, 2), ""])
        block.append(["Top 3 giảm mạnh nhất:", "", "", ""])
        for addr, diff in s["top_losers"]:
            block.append(["", addr, round(diff, 2), ""])
        block.append(["", "", "", ""])  # dòng trống ngăn cách

        ws.update(f"A{row}", block, value_input_option="USER_ENTERED")
        format_cell_range(ws, f"A{row}", CellFormat(textFormat=TextFormat(bold=True)))
        row += len(block)

    # Bảng lịch sử để vẽ biểu đồ xu hướng (mỗi lần chạy thêm 1 dòng)
    history_header_row = row + 1
    values = ws.get_all_values()
    has_history_header = (
        len(values) >= history_header_row and
        len(values[history_header_row - 1]) >= 1 and
        values[history_header_row - 1][0] == "Ngày"
    )
    if not has_history_header:
        header = ["Ngày"] + [f"Tổng cung {t}" for t in summaries] + [f"Số ví {t}" for t in summaries]
        ws.update(f"A{history_header_row}", [header])
        format_cell_range(ws, f"A{history_header_row}:{chr(64+len(header))}{history_header_row}", HEADER_FMT)

    history_row = [date_label] + [round(s["total_today"], 2) for s in summaries.values()] + \
                  [s["wallet_count"] for s in summaries.values()]
    ws.append_rows([history_row], value_input_option="USER_ENTERED",
                    table_range=f"A{history_header_row}")


def main():
    client = get_client()
    sh = client.open_by_key(os.environ["GOOGLE_SHEETS_ID"])
    date_label = datetime.now(VN_TZ).strftime("%d/%m/%Y %H:%M")

    summaries = {}
    for token_name, jetton_address in TOKENS.items():
        print(f"[{token_name}] fetching top holders...")
        snap = fetch_token_snapshot(token_name, jetton_address)
        ws = get_or_create_worksheet(sh, f"{token_name}_Log", rows=20000, cols=6)
        ensure_log_header(ws)
        summary = append_log_rows(ws, date_label, snap["holders"])
        summaries[token_name] = summary
        print(f"[{token_name}] done: {summary['wallet_count']} ví, "
              f"{summary['new_wallets']} mới, {summary['dropped_wallets']} rớt hạng")

    update_dashboard(sh, date_label, summaries)
    print("Đã cập nhật Google Sheets xong.")


if __name__ == "__main__":
    main()

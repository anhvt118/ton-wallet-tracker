"""
Đẩy dữ liệu top holders lên Google Sheets, trình bày dạng BẢNG MA TRẬN
giống hệt bản Excel: mỗi token 1 sheet, mỗi dòng 1 ví, mỗi lần chạy
thêm 1 cột ngày giờ + 1 cột "Chênh lệch" (tô xanh = tăng, đỏ = giảm).

Ngoài ra có thêm sheet "Dashboard" tóm tắt: tổng cung, số ví mới/rớt hạng,
top ví tăng/giảm mạnh nhất mỗi token, + 1 bảng lịch sử nhỏ để cắm biểu đồ
xu hướng.

Cần biến môi trường:
    GOOGLE_SHEETS_CREDENTIALS  - nội dung JSON của service account
    GOOGLE_SHEETS_ID           - ID của Google Sheet
    TONAPI_KEY                 - (tùy chọn) API key TonAPI

Chạy thủ công:
    python push_to_gsheets.py
"""

import os
import json
import string
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

HEADER_BG = Color(0.12, 0.31, 0.47)
HEADER_FMT = CellFormat(
    backgroundColor=HEADER_BG,
    textFormat=TextFormat(bold=True, foregroundColor=Color(1, 1, 1)),
    horizontalAlignment="CENTER",
)
GREEN = Color(0.78, 0.94, 0.81)
RED = Color(0.98, 0.78, 0.81)


def col_letter(idx0):
    letters = ""
    n = idx0 + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters


def get_client():
    info = json.loads(os.environ["GOOGLE_SHEETS_CREDENTIALS"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_or_create_worksheet(sh, title, rows=300, cols=26):
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def update_token_sheet_matrix(ws, token_name, holders, timestamp_label):
    values = ws.get_all_values()
    if not values:
        header = ["Địa chỉ ví", "Token"]
        values = [header]
    else:
        header = values[0]

    has_diff_last = len(header) >= 3 and header[-1] == "Chênh lệch"
    prev_date_idx = (len(header) - 2) if has_diff_last else (len(header) - 1 if len(header) >= 3 else None)

    addr_row = {}
    for i, row in enumerate(values[1:], start=1):
        if row and row[0]:
            addr_row[row[0]] = i

    today_balances = dict(holders)
    all_addrs = dict(today_balances)
    for addr in addr_row:
        if addr not in all_addrs:
            all_addrs[addr] = 0

    new_date_idx = len(header)
    diff_idx = new_date_idx + 1

    header = header + [timestamp_label, "Chênh lệch"]
    values[0] = header

    def pad(row, length):
        return row + [""] * (length - len(row))

    movers = []
    new_wallets = 0
    dropped_wallets = 0

    next_new_row = len(values)
    for addr, bal in all_addrs.items():
        prev_val = None
        if addr in addr_row:
            r = addr_row[addr]
            values[r] = pad(values[r], new_date_idx)
            if prev_date_idx is not None and len(values[r]) > prev_date_idx and values[r][prev_date_idx] not in ("", None):
                try:
                    prev_val = float(str(values[r][prev_date_idx]).replace(",", ""))
                except ValueError:
                    prev_val = None
        else:
            r = next_new_row
            next_new_row += 1
            values.append([addr, token_name])
            addr_row[addr] = r
            values[r] = pad(values[r], new_date_idx)
            new_wallets += 1

        diff = (bal - prev_val) if prev_val is not None else ""
        if prev_val is not None and bal == 0 and prev_val > 0:
            dropped_wallets += 1
        row = values[r]
        row = pad(row, new_date_idx)
        row.append(round(bal, 4))
        row.append(round(diff, 4) if diff != "" else "")
        values[r] = row
        if diff != "":
            movers.append((addr, diff))

    final_len = len(header)
    for i in range(len(values)):
        values[i] = pad(values[i], final_len)

    ws.resize(rows=max(len(values) + 20, ws.row_count), cols=max(final_len + 2, ws.col_count))
    ws.update("A1", values, value_input_option="USER_ENTERED")

    c1, c2 = col_letter(new_date_idx), col_letter(diff_idx)
    format_cell_range(ws, f"{c1}1:{c2}1", HEADER_FMT)
    if new_date_idx == 2:
        set_frozen(ws, rows=1, cols=2)
        format_cell_range(ws, "A1:B1", HEADER_FMT)

    rules = get_conditional_format_rules(ws)
    rules.append(ConditionalFormatRule(
        ranges=[GridRange(sheetId=ws.id, startRowIndex=1, endRowIndex=20000,
                           startColumnIndex=diff_idx, endColumnIndex=diff_idx + 1)],
        booleanRule=BooleanRule(condition=BooleanCondition("NUMBER_GREATER", ["0"]),
                                 format=CellFormat(backgroundColor=GREEN)),
    ))
    rules.append(ConditionalFormatRule(
        ranges=[GridRange(sheetId=ws.id, startRowIndex=1, endRowIndex=20000,
                           startColumnIndex=diff_idx, endColumnIndex=diff_idx + 1)],
        booleanRule=BooleanRule(condition=BooleanCondition("NUMBER_LESS", ["0"]),
                                 format=CellFormat(backgroundColor=RED)),
    ))
    rules.save()

    movers.sort(key=lambda x: x[1], reverse=True)
    top_gainers = movers[:3]
    top_losers = sorted(movers, key=lambda x: x[1])[:3]
    total_today = sum(today_balances.values())
    total_prev = None
    if prev_date_idx is not None:
        total_prev = 0.0
        for row in values[1:]:
            v = row[prev_date_idx] if len(row) > prev_date_idx else ""
            if v not in ("", None):
                try:
                    total_prev += float(str(v).replace(",", ""))
                except ValueError:
                    pass

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
        block.append(["", "", "", ""])

        ws.update(f"A{row}", block, value_input_option="USER_ENTERED")
        format_cell_range(ws, f"A{row}", CellFormat(textFormat=TextFormat(bold=True)))
        row += len(block)

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
        format_cell_range(ws, f"A{history_header_row}:{col_letter(len(header)-1)}{history_header_row}", HEADER_FMT)

    history_row = [date_label] + [round(s["total_today"], 2) for s in summaries.values()] + \
                  [s["wallet_count"] for s in summaries.values()]
    ws.append_rows([history_row], value_input_option="USER_ENTERED",
                    table_range=f"A{history_header_row}")


def main():
    client = get_client()
    sh = client.open_by_key(os.environ["GOOGLE_SHEETS_ID"])
    date_label = datetime.now(VN_TZ).strftime("%d/%m/%Y")

    summaries = {}
    for token_name, jetton_address in TOKENS.items():
        print(f"[{token_name}] fetching top holders...")
        snap = fetch_token_snapshot(token_name, jetton_address)
        ws = get_or_create_worksheet(sh, token_name, rows=300, cols=30)
        summary = update_token_sheet_matrix(ws, token_name, snap["holders"], date_label)
        summaries[token_name] = summary
        print(f"[{token_name}] done: {summary['wallet_count']} ví, "
              f"{summary['new_wallets']} mới, {summary['dropped_wallets']} rớt hạng")

    update_dashboard(sh, date_label, summaries)
    print("Đã cập nhật Google Sheets xong.")


if __name__ == "__main__":
    main()

"""
Module dùng chung: gọi TonAPI để lấy thông tin jetton + top holders.
Được import bởi cả update_tracker.py (ghi ra Excel) và push_to_gsheets.py
(ghi ra Google Sheets) để không lặp code.
"""

import os
import time
import base64
import requests

# Mỗi token: tên hiển thị -> địa chỉ contract (jetton master) trên TON
TOKENS = {
    "DIDI": "EQCRUitj7ehYvSzZKTyhq02-HpbhLNgAvnMF5I7Dx31QxIAH",
    "YODA": "EQC7vuKEYLdC72YhUWt3AUVA-Oi66Q1DxTHXH7r6pXaV50j7",
    "UTYA": "EQBaCgUwOoc6gHCNln_oJzb0mVs79YG7wYoavh-o1ItaneLA",
    "REDO": "EQBZ_cafPyDr5KUTs0aNxh0ZTDhkpEZONmLJA2SNGlLm4Cko",
}

TOP_N = 100  # số ví top đầu muốn theo dõi mỗi token

TONAPI_BASE = "https://tonapi.io"
TONAPI_KEY = os.environ.get("TONAPI_KEY", "")


def tonapi_get(path, params=None):
    headers = {}
    if TONAPI_KEY:
        headers["Authorization"] = f"Bearer {TONAPI_KEY}"
    for attempt in range(5):
        resp = requests.get(f"{TONAPI_BASE}{path}", params=params, headers=headers, timeout=30)
        if resp.status_code == 429:
            time.sleep(3 * (attempt + 1))
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()


def _crc16_xmodem(data: bytes) -> int:
    crc = 0
    for b in data:
        crc ^= b << 8
        for _ in range(8):
            crc <<= 1
            if crc & 0x10000:
                crc ^= 0x1021
            crc &= 0xFFFF
    return crc


def to_friendly_address(raw_address, bounceable=True, testnet=False):
    """Chuyển '0:abcd...' (raw) -> dạng thân thiện EQ.../UQ... (base64url).
    Nếu input đã ở dạng thân thiện sẵn (không có dấu ':') thì trả nguyên."""
    if not raw_address or ":" not in raw_address:
        return raw_address
    try:
        workchain_str, hash_hex = raw_address.split(":", 1)
        workchain = int(workchain_str)
        hash_bytes = bytes.fromhex(hash_hex)
        if len(hash_bytes) != 32:
            return raw_address
    except (ValueError, TypeError):
        return raw_address

    tag = 0x11 if bounceable else 0x51
    if testnet:
        tag |= 0x80

    buf = bytes([tag]) + workchain.to_bytes(1, "big", signed=True) + hash_bytes
    crc = _crc16_xmodem(buf)
    buf += crc.to_bytes(2, "big")
    return base64.urlsafe_b64encode(buf).decode("ascii")


def get_jetton_decimals(jetton_address):
    info = tonapi_get(f"/v2/jettons/{jetton_address}")
    meta = info.get("metadata", {})
    return int(meta.get("decimals", 9)), meta.get("symbol", "")


def get_top_holders(jetton_address, limit=TOP_N):
    """Trả về list[(address, balance_raw_int)] sắp theo số dư giảm dần.
    Địa chỉ luôn được chuẩn hóa về dạng thân thiện (EQ.../UQ...)."""
    data = tonapi_get(f"/v2/jettons/{jetton_address}/holders", params={"limit": limit, "offset": 0})
    holders = []
    for item in data.get("addresses", []):
        owner = item.get("owner", {}) or {}
        addr = owner.get("address") or item.get("address")
        addr = to_friendly_address(addr)
        balance_raw = item.get("balance", "0")
        holders.append((addr, int(balance_raw)))
    return holders


def fetch_token_snapshot(token_name, jetton_address):
    """Lấy 1 snapshot đầy đủ cho 1 token: decimals, symbol, holders (đã quy đổi số thực)."""
    decimals, symbol = get_jetton_decimals(jetton_address)
    holders_raw = get_top_holders(jetton_address, limit=TOP_N)
    holders = [(addr, bal / (10 ** decimals)) for addr, bal in holders_raw]
    return {"token": token_name, "decimals": decimals, "symbol": symbol, "holders": holders}

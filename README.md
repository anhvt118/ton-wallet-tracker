# Theo dõi ví TON tự động (DIDI, YODA)

Tự động lấy top 100 ví nắm giữ nhiều nhất của mỗi token qua TonAPI, chạy
1 lần/ngày lúc 8h sáng (giờ VN) qua GitHub Actions, ghi ra 2 nơi:

1. **Excel** (`data/theo_doi_vi.xlsx`) — dạng bảng ma trận, backup trong repo.
2. **Google Sheets** — dạng chuyên nghiệp hơn để xem/chia sẻ:
   - Sheet `DIDI_Log`, `YODA_Log`: dạng dài (mỗi dòng = 1 ví + 1 ngày),
     dễ lọc, pivot table, vẽ biểu đồ. Cột "Chênh lệch" tự tô xanh (tăng)
     / đỏ (giảm).
   - Sheet `Dashboard`: tổng cung top 100 mỗi token, chênh lệch so lần
     trước, số ví mới/rớt hạng, top 3 ví tăng/giảm mạnh nhất, + 1 bảng
     lịch sử nhỏ để cắm biểu đồ xu hướng.

## Chạy thử trên máy mình

```bash
pip install -r requirements.txt
export TONAPI_KEY=xxxx                          # optional
export GOOGLE_SHEETS_CREDENTIALS='{"type": ...}' # nội dung JSON service account
export GOOGLE_SHEETS_ID=xxxxxxxxxxxxxxxxxxx      # ID trong URL sheet

python update_tracker.py       # ghi ra Excel
python push_to_gsheets.py      # đẩy lên Google Sheets
```

## Setup Google Sheets (làm 1 lần)

### 1. Tạo Google Sheet
Tạo 1 Google Sheet trống, đặt tên tùy ý. Copy **ID** của nó từ URL:
`https://docs.google.com/spreadsheets/d/`**`ID_Ở_ĐÂY`**`/edit`

### 2. Tạo service account (tài khoản "robot" để script ghi vào sheet)
1. Vào [console.cloud.google.com](https://console.cloud.google.com), tạo 1 project mới (hoặc dùng project có sẵn).
2. Vào **APIs & Services → Library**, bật 2 API: **Google Sheets API** và **Google Drive API**.
3. Vào **APIs & Services → Credentials → Create Credentials → Service account**, đặt tên tùy ý, bấm qua các bước còn lại (không cần cấp quyền gì thêm).
4. Vào service account vừa tạo → tab **Keys → Add key → Create new key → JSON** → tải file JSON về.
5. Mở file JSON đó, copy **địa chỉ email** (dạng `xxx@xxx.iam.gserviceaccount.com`).

### 3. Chia sẻ Sheet cho service account
Mở Google Sheet đã tạo ở bước 1 → **Share** → dán email service account vào → chọn quyền **Editor** → **Send** (bỏ qua cảnh báo không gửi được thông báo).

### 4. Thêm secrets vào GitHub
Vào repo → **Settings → Secrets and variables → Actions → New repository secret**, tạo 2 secret:
- `GOOGLE_SHEETS_CREDENTIALS`: mở file JSON tải ở bước 2.4, copy **toàn bộ nội dung**, dán vào.
- `GOOGLE_SHEETS_ID`: dán ID lấy ở bước 1.

(Secret `TONAPI_KEY` nếu đã có từ trước thì giữ nguyên, không cần đổi.)

### 5. Chạy thử
Vào tab **Actions → Update wallet tracker → Run workflow**. Sau khi chạy
xong (dấu tích xanh), mở Google Sheet lên sẽ thấy sheet `DIDI_Log`,
`YODA_Log`, `Dashboard` xuất hiện.

### 6. Cắm biểu đồ xu hướng (làm 1 lần, thủ công trên Google Sheets)
1. Mở sheet `Dashboard`, kéo xuống tìm bảng có header **"Ngày | Tổng cung DIDI | Tổng cung YODA | ..."**.
2. Bôi đen **toàn bộ cột** của bảng đó (bấm vào chữ cột, ví dụ kéo từ cột A tới cột E) — bôi cả cột chứ không chỉ vùng có dữ liệu, để biểu đồ tự nhận dòng mới thêm vào mỗi ngày mà không cần sửa lại.
3. **Insert → Chart**. Google Sheets sẽ tự đề xuất biểu đồ đường (Line chart) — chỉnh sửa tiêu đề, màu sắc tùy ý trong panel **Customize**.
4. Từ hôm sau, script tự thêm dòng mới vào bảng lịch sử này, biểu đồ tự cập nhật, không cần làm lại bước này nữa.

Muốn có thêm biểu đồ khác (vd. so sánh top 10 ví lớn nhất) thì làm tương
tự: bôi đen vùng dữ liệu trong `DIDI_Log`/`YODA_Log` (lọc trước bằng
Data → Create a filter nếu cần) rồi Insert → Chart.

## Thêm/bớt token theo dõi

Sửa dict `TOKENS` trong `ton_data.py`:

```python
TOKENS = {
    "DIDI": "EQCRUitj7ehYvSzZKTyhq02-HpbhLNgAvnMF5I7Dx31QxIAH",
    "YODA": "EQC7vuKEYLdC72YhUWt3AUVA-Oi66Q1DxTHXH7r6pXaV50j7",
    # "TENTOKEN": "địa_chỉ_contract_jetton",
}
```

Token mới sẽ tự có sheet `<TENTOKEN>_Log` riêng và tự gộp vào `Dashboard`.

## Chi phí

- TonAPI: free tier đủ dùng (1 lần/ngày, 2-3 token).
- Google Sheets API: miễn phí hoàn toàn cho mức sử dụng này.
- GitHub Actions: free tier private repo ~2000 phút/tháng, job này chạy
  vài chục giây/lần, 1 lần/ngày — không tốn phí.

## Lưu ý

- Ví rớt khỏi top 100 ở lần chạy sau -> ghi số dư 0 cho ngày đó (Log) /
  cột ngày đó (Excel).
- Ví mới lọt top 100 -> tự thêm dòng, cột Chênh lệch để trống lần đầu
  (chưa có gì để so sánh).

# Theo dõi ví TON tự động (DIDI, YODA)

Script này thay thế việc chụp ảnh thủ công: tự gọi API công khai của TON
(`tonapi.io`) để lấy top 100 ví nắm giữ nhiều nhất của mỗi token, rồi cập
nhật vào file Excel `data/theo_doi_vi.xlsx` — mỗi token 1 sheet, mỗi lần
chạy thêm 1 cột thời gian + 1 cột chênh lệch (xanh = tăng, đỏ = giảm),
đúng format đang dùng.

## Chạy thử trên máy mình (không cần GitHub)

```bash
pip install -r requirements.txt
python update_tracker.py
```

File kết quả nằm ở `data/theo_doi_vi.xlsx`. Muốn chạy lại nhiều lần trong
ngày để test thì cứ chạy lệnh trên nhiều lần, mỗi lần sẽ thêm 1 cột mới.

## Tự động chạy mỗi 12 tiếng — dùng GitHub Actions (miễn phí)

1. Tạo 1 repo GitHub mới (private cũng được), đẩy toàn bộ thư mục này lên.
2. Lấy API key miễn phí tại https://tonconsole.com (mục TonAPI) để tránh
   bị giới hạn tốc độ — không bắt buộc nhưng nên có.
3. Trong repo trên GitHub: vào **Settings → Secrets and variables →
   Actions → New repository secret**, đặt tên `TONAPI_KEY`, dán key vào.
4. Xong — workflow trong `.github/workflows/update.yml` sẽ tự chạy lúc
   7h sáng và 7h tối (giờ VN) mỗi ngày, tự cập nhật file Excel và commit
   lại vào repo. Muốn đổi giờ thì sửa dòng `cron` trong file đó.
5. Muốn chạy tay ngay lập tức để test: vào tab **Actions** trên GitHub →
   chọn workflow "Update wallet tracker" → **Run workflow**.

Xem file mới nhất: mở `data/theo_doi_vi.xlsx` trong repo (GitHub cho xem
trực tiếp), hoặc pull về máy / tải xuống.

## Thêm/bớt token theo dõi

Sửa dict `TOKENS` ở đầu file `update_tracker.py`:

```python
TOKENS = {
    "DIDI": "EQCRUitj7ehYvSzZKTyhq02-HpbhLNgAvnMF5I7Dx31QxIAH",
    "YODA": "EQC7vuKEYLdC72YhUWt3AUVA-Oi66Q1DxTHXH7r6pXaV50j7",
    # "TENTOKEN": "địa_chỉ_contract_jetton",
}
```

## Lưu ý

- Free tier của TonAPI đủ dùng cho lịch chạy 12h/lần với vài token.
- Nếu 1 ví rớt khỏi top 100 ở lần chạy sau, script tự set số dư ngày đó
  về 0 (giống cách đang làm thủ công).
- Nếu 1 ví mới lọt vào top 100, script tự thêm dòng mới, cột chênh lệch
  để trống cho lần đầu (chưa có dữ liệu để so sánh).

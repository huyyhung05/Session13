# GreenTech K26 Infrastructure

**Repository:** `CNTT1_NhapMonCNTT_MiniProject_GreenTech`  
**Sinh viên:** Phan Huy Hưng — CNTT1 — Nhập môn CNTT

## 1. Kiến trúc 4 trụ cột

```text
CNTT1_NhapMonCNTT_MiniProject_GreenTech/
├── docs/
│   ├── README.md
│   ├── srs-he-thong.docx
│   └── bao-cao-phan-tich.docx
├── scripts/
│   ├── auto_report.bat
│   └── setup-command.txt
├── data/
│   ├── log-he-thong.log
│   └── quan-ly-thiet-bi.xlsx
├── backup/
│   └── bao-cao-su-co.txt
├── .env
└── .gitignore
```

## 2. IPO

| Thành phần | Nội dung |
|---|---|
| Input | `data/log-he-thong.log`, mã linh kiện, dữ liệu 100 trạm |
| Process | Batch + Pipe + `findstr`; Excel `TRIM` + `VLOOKUP` |
| Output | `backup/bao-cao-su-co.txt`, bảng trạng thái thiết bị đã đối soát |

## 3. DIKW

- **Data:** log thô và mã thiết bị.
- **Information:** các dòng `ERROR`, `CRITICAL`, `TIMEOUT` đã được lọc.
- **Knowledge:** nhận ra trạm hoặc thiết bị có lỗi lặp lại.
- **Wisdom:** quyết định bảo trì, thay storage hoặc nâng cấp cấu hình.

## 4. Khởi tạo repository bằng 1 dòng

```cmd
mkdir CNTT1_NhapMonCNTT_MiniProject_GreenTech\docs CNTT1_NhapMonCNTT_MiniProject_GreenTech\scripts CNTT1_NhapMonCNTT_MiniProject_GreenTech\data CNTT1_NhapMonCNTT_MiniProject_GreenTech\backup
```

Tạo `.env` bằng Relative Path sau khi `cd CNTT1_NhapMonCNTT_MiniProject_GreenTech`:

```cmd
echo GT_DB_PASSWORD=demo-password> .env
attrib +h .env
```

> `.env` chỉ chứa cấu hình giả định trong bài và đã được thêm vào `.gitignore`.

## 5. Chạy Robot lọc log

```cmd
scripts\auto_report.bat
```

Script đọc log theo luồng và lọc `CRITICAL`, `ERROR`, `TIMEOUT`; kết quả được **ghi nối** bằng `>>`, không ghi đè báo cáo cũ.

## 6. Checklist vận hành

- [x] Kiểm tra cây thư mục không có dấu/khoảng trắng
- [x] Kiểm tra `.env` không commit công khai
- [x] Kiểm tra log có thể lọc bằng Batch
- [x] Kiểm tra Excel không Merge Cell / hàng trống
- [x] Kiểm tra `TRIM` xử lý mã có khoảng trắng
- [x] Kiểm tra `VLOOKUP(...,FALSE)` trả đúng tên thiết bị
- [ ] Thay dữ liệu mẫu bằng dữ liệu thật trước khi vận hành thực tế

## 7. Thông số phần cứng mẫu

| Mã linh kiện | Thiết bị | Storage | RAM tham khảo |
|---|---|---:|---:|
| HW001 | Gateway IoT | SSD | 4 GB |
| HW006 | Edge Server | SSD | 16 GB |
| HW007 | Backup Node | HDD | 8 GB |

## 8. Công thức Excel

```text
=TRIM(B2)
=IFERROR(VLOOKUP(C2,'danh-muc-linh-kien'!$A$2:$D$9,2,FALSE),"Chưa cập nhật")
```

## 9. Kiểm thử tốc độ mạng

Công thức: `MB/s = Mbps / 8`. Ví dụ mạng **100 Mbps** tương đương **12,5 MB/s**. Tệp log **500 MB** cần tối thiểu khoảng **40 giây** trong điều kiện lý tưởng.

## 10. Audit trước khi chạy lệnh AI

1. Đọc và hiểu lệnh.
2. Kiểm tra phạm vi đường dẫn.
3. Chạy thử trên dữ liệu mẫu/backup.

**Không chạy** lệnh xóa hàng loạt như `del /f /s /q` nếu chưa xác nhận chính xác đường dẫn và hậu quả.

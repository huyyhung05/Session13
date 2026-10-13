**Đề bài / Yêu cầu**

## **1. Tổng quan**

### **1.1. Mục đích**

Nội dung giúp sinh viên:

- Hiểu mô hình xử lý thông tin IPO và tư duy hệ thống DIKW để tổ chức dữ liệu khoa học.
- Thiết kế được kiến trúc lưu trữ dự án theo mô hình cây thư mục chuẩn hóa.
- Sử dụng Terminal, Shell Script và toán tử Pipe để xử lý tệp nhật ký dung lượng lớn.
- Xây dựng tài liệu kỹ thuật chuyên nghiệp bằng Markdown và Word Styles.
- Vận dụng tư duy Database Mindset và hàm VLOOKUP để đối soát dữ liệu trong Excel.

### **1.2. Công nghệ sử dụng**

- Command Line Interface (CMD / Terminal / Shell).
- Shell Scripting (.bat / .sh).
- Markdown.
- Microsoft Word (Heading Styles, Auto TOC).
- Microsoft Excel (VLOOKUP, TRIM, AutoFilter).
- Git & GitHub.

# **PHẦN 1. TƯ DUY HỆ THỐNG VÀ KIẾN TRÚC LƯU TRỮ**

## **2. Mô hình xử lý thông tin IPO và DIKW**

### **Mục tiêu**

Sinh viên hiểu được:

- Vì sao một hệ thống hạ tầng cần được nhìn theo luồng xử lý thông tin.
- IPO là gì và áp dụng vào bài toán quản lý log.
- DIKW là gì và áp dụng vào việc chuyển dữ liệu thô thành quyết định vận hành.

### **Nội dung chính**

**Mô hình IPO (Input – Process – Output):**

- Input: Tệp log-he-thong.log phát sinh liên tục từ máy chủ.
- Process: Script auto_report.bat lọc các dòng CRITICAL / ERROR bằng Pipe.
- Output: Tệp bao-cao-su-co.txt trong thư mục backup/ dùng để báo cáo.

**Mô hình DIKW (Data – Information – Knowledge – Wisdom):**

- Data: Hàng triệu dòng log thô, chưa có ý nghĩa khi đứng riêng lẻ.
- Information: Các dòng log đã được lọc theo từ khóa lỗi (ERROR 404, TIMEOUT).
- Knowledge: Nhận biết mẫu lỗi lặp lại, trạm nào có tần suất sự cố cao.
- Wisdom: Ra quyết định vận hành — nâng cấp thiết bị, điều chỉnh cấu hình mạng.

## **3. Kiến trúc cây thư mục 4 trụ cột**

### **Khái niệm**

Một dự án hạ tầng chuyên nghiệp cần tách biệt rõ 4 nhóm nội dung (4 trụ cột), tránh việc trộn lẫn mã nguồn, dữ liệu và tài liệu trong cùng một thư mục.

- docs/ — Trụ cột Tài liệu: README, đặc tả kỹ thuật.
- scripts/ — Trụ cột Xử lý: các tệp tự động hóa .bat/.sh.
- data/ — Trụ cột Dữ liệu: log hệ thống, bảng đối soát thiết bị.
- backup/ — Trụ cột Lưu trữ: bản sao báo cáo, phục vụ truy vết.

### **Thiết kế cây thư mục**

| greentech-system/<br>├── docs/<br>│   ├── [README.md](http://readme.md/)<br>│   └── srs-he-thong.docx<br>├── scripts/<br>│   └── auto_report.bat<br>├── data/<br>│   ├── log-he-thong.log<br>│   └── quan-ly-thiet-bi.xlsx<br>├── backup/<br>│   └── bao-cao-su-co.txt<br>└── .env |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |

### **Quy tắc đặt tên**

- Sử dụng kebab-case cho tên thư mục và tệp (VD: log-he-thong.log).
- Không sử dụng Tiếng Việt có dấu.
- Không sử dụng khoảng trắng trong tên thư mục/tệp.
- Tệp cấu hình nhạy cảm (.env) phải là tệp ẩn, không commit lên GitHub công khai.

## **4. Thiết kế cơ sở dữ liệu Excel (Asset Database)**

### **Khái niệm**

Bảng tính Excel cần một trường định danh duy nhất (khóa chính) ở bảng danh mục, và một trường tham chiếu (khóa ngoại) ở bảng còn lại để tránh lặp dữ liệu.

- Bảng cha: danh-muc-linh-kien — nơi lưu thông tin gốc của từng loại linh kiện.
- Bảng con: trang-thai-tram — mỗi trạm tham chiếu đến một mã linh kiện.

### **Thiết kế bảng**

**Sheet danh-muc-linh-kien**

- ma_linh_kien (Khóa chính)
- ten_thiet_bi
- don_gia
- loai_storage (SSD/HDD)

**Sheet trang-thai-tram**

- ma_tram
- ma_linh_kien (Khóa ngoại)
- dung_luong_ram
- ngay_lap_dat

Quan hệ: [trang-thai-tram.ma](http://trang-thai-tram.ma/)\_linh_kien → [danh-muc-linh-kien.ma](http://danh-muc-linh-kien.ma/)\_linh_kien

### **Quy tắc chính**

- Khóa ngoại (ma_linh_kien) đặt ở bảng phía nhiều — trang-thai-tram.
- Không để trống ma_linh_kien khi trạm đã lắp thiết bị.
- Vùng dữ liệu phải liền mạch, không có hàng trống hoặc ô Merge Cell.
- Mã linh kiện phải được chuẩn hóa (loại bỏ khoảng trắng thừa) trước khi đối chiếu.

# **PHẦN 2. XÂY DỰNG CÁC MODULE TỰ ĐỘNG HÓA**

## **5. Bối cảnh**

Xây dựng hệ thống quản trị hạ tầng gồm 4 module chức năng, tương ứng với 4 trụ cột của cây thư mục.

| **Module**                  | **Công cụ chính**                | **Thư mục liên quan**    |
| :-------------------------- | :------------------------------- | :----------------------- |
| Khởi tạo cấu trúc dự án     | Terminal / CMD                   | toàn bộ cây thư mục      |
| Robot lọc Log tự động       | Shell Script, Pipe, findstr/grep | scripts/, data/, backup/ |
| Tài liệu hướng dẫn vận hành | Markdown, Word Styles            | docs/                    |
| Đối soát tài nguyên         | Excel VLOOKUP, TRIM              | data/                    |

## **6. Module 1: Khởi tạo cấu trúc dự án**

### **Chức năng**

Tạo toàn bộ cây thư mục dự án chỉ bằng một dòng lệnh duy nhất, sử dụng cú pháp lồng nhau.

### **Quy trình xử lý**

1. Xác định vị trí thư mục gốc cần tạo (Root).
2. Soạn một dòng lệnh mkdir duy nhất, liệt kê toàn bộ thư mục con lồng nhau.
3. Thực thi lệnh, kiểm tra cây thư mục vừa tạo bằng lệnh tree.
4. Di chuyển vào từng thư mục con bằng Relative Path để tạo tệp cấu hình .env.
5. Xác nhận .env là tệp ẩn và không bị đẩy lên Git khi chưa được whitelist.

### **Lệnh mẫu**

| mkdir greentech-system\docs greentech-system\scripts ^<br>  greentech-system\data greentech-system\backup<br><br><br>cd greentech-system<br>echo. > .env |
| -------------------------------------------------------------------------------------------------------------------------------------------------------- |

### **Lưu ý / Bẫy**

- Tên thư mục tuyệt đối không dùng Tiếng Việt có dấu hoặc khoảng trắng.
- Không dùng đường dẫn tuyệt đối cố định (VD: C:\Users\Ten) để đảm bảo script chạy được trên máy khác.

## **7. Module 2: Robot lọc Log tự động**

Script: scripts/auto_report.bat

### **Chức năng**

Tự động quét tệp log-he-thong.log (có thể lên tới 500MB), lọc ra các dòng nghiêm trọng và ghi nối kết quả vào báo cáo, mà không cần mở toàn bộ tệp trong RAM.

### **Quy trình xử lý**

6. Dùng lệnh echo để in tiêu đề báo cáo (kèm ngày giờ chạy).
7. Dùng findstr (Windows) hoặc grep (Linux/macOS) để tìm các dòng chứa "CRITICAL" hoặc "ERROR".
8. Kết hợp toán tử Pipe (|) để dòng dữ liệu chảy trực tiếp từ lệnh đọc file sang lệnh lọc.
9. Dùng toán tử >> để nối kết quả vào cuối tệp bao-cao-su-co.txt, tránh ghi đè báo cáo cũ.
10. Lưu tệp báo cáo vào thư mục backup/.

### **Script mẫu (auto_report.bat)**

| @echo off<br>echo ===== BAO CAO SU CO HE THONG ===== >> backup\bao-cao-su-co.txt<br>echo Thoi gian chay: %date% %time% >> backup\bao-cao-su-co.txt<br><br><br>type data\log-he-thong.log \| findstr /I "CRITICAL ERROR TIMEOUT" >> backup\bao-cao-su-co.txt<br><br><br>echo Hoan tat loc log. |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |

### **Output mẫu (backup/bao-cao-su-co.txt)**

| ===== BAO CAO SU CO HE THONG =====<br>Thoi gian chay: Mon 08/03/2026 10:00:00<br>[10:00:12] ERROR 404 - tram-cam-bien-27 khong phan hoi<br>[10:00:45] CRITICAL - mat ket noi may chu trung tam<br>[10:01:03] TIMEOUT - tram-cam-bien-58 qua thoi gian cho |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |

### **Lưu ý / Bẫy**

- Dùng nhầm toán tử > thay vì >> sẽ ghi đè và xóa mất báo cáo cũ.
- Phải có bước Audit (thẩm định) trước khi chạy bất kỳ lệnh xóa nào do AI gợi ý, tuyệt đối tránh del /f /s /q khi chưa xác nhận đường dẫn.

## **8. Module 3: Tài liệu hướng dẫn vận hành**

### **Chức năng**

Chuẩn hóa tài liệu nội bộ để nhân viên mới có thể vận hành hệ thống mà không cần hỏi lại kỹ sư phụ trách.

### **Quy trình xử lý — [README.md](http://readme.md/)**

11. Đặt tên hệ thống ở cấp Heading 1 (#).
12. Viết hướng dẫn chạy script trong khối Code Block.
13. Tạo Checklist kiểm tra thiết bị bằng Task List (- [ ]).
14. Trình bày thông số phần cứng bằng bảng Markdown (Table).

### [**README.md**](http://readme.md/) **mẫu**

| # He thong quan tri ha tang GreenTech K26<br><br><br>## Cach chay script loc log<br>cd scripts<br>auto_report.bat<br><br><br>## Checklist kiem tra thiet bi<br>- [ ] Kiem tra ket noi tram cam bien<br>- [ ] Kiem tra dung luong RAM con trong<br>- [ ] Sao luu bao cao vao backup/<br><br><br>\| Tram \| RAM \| Ngay lap dat \|<br>\|------\|-----\|--------------\|<br>\| tram-01 \| 4GB \| 2026-01-10 \| |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |

### **Quy trình xử lý — SRS Document (Word)**

15. Gán nhãn Heading Styles (1, 2, 3) cho toàn bộ tiêu đề mục lục.
16. Chèn mục lục tự động (Auto TOC) ngay sau trang bìa.
17. Kiểm tra TOC cập nhật đúng số trang trước khi xuất bản.

## **9. Module 4: Đối soát tài nguyên (Excel)**

### **Chức năng**

Tự động điền tên thiết bị vào bảng trạng thái trạm dựa trên mã linh kiện, tránh nhập tay gây sai lệch dữ liệu.

### **Quy trình xử lý**

18. Làm sạch dữ liệu: dùng hàm TRIM để loại bỏ khoảng trắng thừa trong cột ma_linh_kien.
19. Xác định vùng dữ liệu tra cứu (bảng danh-muc-linh-kien) và cố định vùng bằng dấu $.
20. Viết công thức VLOOKUP với tham số FALSE để chỉ nhận kết quả trùng khớp chính xác.
21. Kiểm tra lỗi #N/A — nếu có, quay lại bước làm sạch dữ liệu.
22. Áp dụng AutoFilter để kiểm tra nhanh theo từng trạm.

### **Công thức mẫu**

| =TRIM(A2)<br>=VLOOKUP(TRIM(B2), 'danh-muc-linh-kien'!$A$2:$D$100, 2, FALSE) |
| --------------------------------------------------------------------------- |

### **Lưu ý / Bẫy**

- Vùng dữ liệu chứa hàng trống hoặc ô Merge Cell sẽ làm AutoFilter hoạt động sai.
- Không cố định vùng dò tìm bằng $ sẽ khiến công thức lệch khi sao chép xuống các dòng khác.

## **10. Các bẫy dữ liệu cần xử lý**

| **Bẫy**                              | **Rủi ro nếu không xử lý**                           | **Cách phòng tránh**                    |
| :----------------------------------- | :--------------------------------------------------- | :-------------------------------------- |
| Tên thư mục có dấu / khoảng trắng    | Script và đường dẫn chạy lỗi trên nhiều hệ điều hành | Luôn đặt tên theo kebab-case, không dấu |
| Lệnh xóa nguy hiểm (del /f /s /q)    | Xóa nhầm toàn bộ dữ liệu dự án                       | Bắt buộc bước Audit trước khi thực thi  |
| Hàng trống / Merge Cell trong Excel  | AutoFilter và VLOOKUP trả kết quả sai                | Chuẩn hóa vùng dữ liệu liền mạch        |
| Khoảng trắng thừa trong mã linh kiện | VLOOKUP trả lỗi #N/A                                 | Dùng hàm TRIM trước khi đối chiếu       |
| Dùng > thay vì >> khi ghi log        | Mất toàn bộ báo cáo cũ                               | Luôn dùng >> khi cần nối thêm dữ liệu   |
| Tính sai đơn vị tốc độ mạng          | Ước lượng sai thời gian tải log                      | Quy đổi Mbps sang MB/s theo hệ số 8     |
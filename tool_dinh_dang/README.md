# JSLHU Auto-Formatter & Reviewer 📝

Một giải pháp tự động hóa toàn diện viết bằng Python, tích hợp trí tuệ nhân tạo (Gemini AI), giúp rà soát chất lượng học thuật và định dạng các bản thảo bài báo khoa học tuân thủ 100% quy chế thể lệ của **Tạp chí Khoa học Lạc Hồng (JSLHU)**.

Dự án cung cấp cả giao diện dòng lệnh (CLI) tiện lợi và giao diện Web (Streamlit Dashboard) cao cấp, thân thiện cho mọi nhà nghiên cứu.

---

## 🌟 Các Tính Năng Vượt Trội

### 1. Auto-Formatter (Định dạng tự động chuẩn LHU)
*   **Lề trang (Page Margins):** Cấu hình chính xác lề trên/dưới `1.5 cm`, trái `2.5 cm`, phải `1.0 cm` đồng bộ toàn bài báo.
*   **Bố cục hai cột (Two Columns):** Tự động chia đều 2 cột với độ rộng mỗi cột là `8.5 cm` và khoảng cách giữa 2 cột là `0.51 cm` (xóa triệt để các thuộc tính đè lề cột cũ trong bản gốc).
*   **Tiêu đề mục (Headings):** Căn lề trái sát mép cột, in đậm, cỡ chữ `10 pt`, tô màu xanh lam thương hiệu LHU (`#365F91`) bắt mắt.
*   **Văn bản thường (Body Text):** Cỡ chữ `10 pt` Times New Roman, thụt đầu dòng đầu tiên `0.36 cm`, khoảng cách sau đoạn `6 pt`, giãn dòng đơn (`Single`).
*   **Căn chỉnh công thức toán học (Equations):** Tự động căn giữa công thức (Center Tab tại `4.25 cm`) và đánh số thứ tự công thức sát lề phải cột (Right Tab tại `8.5 cm`). Đảm bảo cỡ chữ công thức toán thô luôn hiển thị đồng bộ `10 pt` Cambria Math.
*   **Tài liệu tham khảo (References):** 
    *   Hỗ trợ đầy đủ định dạng **LNCS (Lecture Notes in Computer Science)** của Springer.
    *   Tự động phát hiện và đổi số thứ tự sang dạng ngoặc vuông `[1]`, `[2]` thay thế cho kiểu chấm.
    *   Xóa sạch các lỗi gạch chân màu xanh do copy liên kết từ phần mềm quản lý trích dẫn.
    *   Áp dụng lùi treo (Hanging Indent) **`0.5 cm`** căn đều hai bên cực kỳ thẳng hàng, đẹp mắt.

### 2. AI Reviewer (Phản biện học thuật thông minh)
*   Tự động đọc nội dung bản thảo và rà soát các chỉ số quy định của Tòa soạn JSLHU:
    *   Độ dài tiêu đề chính (không quá 20 từ).
    *   Độ dài tóm tắt tiếng Việt và tiếng Anh (từ 150 - 250 từ).
    *   Số lượng từ khóa (từ 3 - 5 từ khóa).
    *   Số lượng tài liệu tham khảo (từ 10 - 15 tài liệu) và kiểm tra tỷ lệ tài liệu mới xuất bản trong 5 năm gần đây (tối thiểu 3 tài liệu).
    *   Kiểm tra email liên hệ của tác giả chính phải có tên miền `@lhu.edu.vn`.
*   Tự động chấm điểm phản biện sơ loại và gợi ý chỉnh sửa chi tiết nội dung khoa học bằng Gemini AI.

---

## 📁 Cấu Trúc Thư Mục Dự Án

```text
├── tool_dinh_dang/
│   ├── app.py                  # Giao diện Web Streamlit cao cấp
│   ├── config.py               # Các cấu hình quy chuẩn và biến môi trường
│   ├── extractor.py            # Hàm trích xuất dữ liệu, cấu trúc tài liệu Word
│   ├── formatter.py            # Core engine xử lý XML định dạng Word
│   ├── reviewer.py             # Module AI rà soát chất lượng học thuật
│   ├── run_for_new_file.py     # Script CLI định dạng nhanh
│   └── BM01. jslhu template.docx # Tệp mẫu Word gốc của LHU
├── requirements.txt            # Danh sách thư viện phụ thuộc
├── .gitignore                  # Ngăn đẩy file rác/key bí mật lên GitHub
└── README.md                   # Hướng dẫn sử dụng dự án
```

---

## 🛠️ Hướng Dẫn Cài Đặt

### 1. Tải mã nguồn về máy
```bash
git clone https://github.com/username/jslhu-formatter.git
cd jslhu-formatter
```

### 2. Cài đặt các thư viện phụ thuộc
Yêu cầu Python từ bản **3.9** trở lên.
```bash
pip install -r requirements.txt
```

### 3. Cấu hình khóa bảo mật API
Tạo tệp `.env` tại thư mục gốc của dự án và điền Gemini API Key của bạn vào:
```env
GEMINI_API_KEY=your_gemini_api_key_here
```
*(Hãy yên tâm, tệp `.env` đã được cấu hình trong `.gitignore` để không bị lộ khi đẩy lên GitHub).*

---

## 🚀 Hướng Dẫn Sử Dụng

### Cách 1: Sử dụng giao diện Web Streamlit (Khuyên dùng)
Khởi chạy ứng dụng web local:
```bash
streamlit run tool_dinh_dang/app.py
```
*   **Trình duyệt mở ra:** Truy cập địa chỉ `http://localhost:8501`.
*   **Sử dụng:** Chọn tab **Định dạng tự động (Auto-Formatter)**, tải lên tệp Word bản thảo thô của bạn (nội dung bắt đầu từ mục *1. Giới thiệu*), điều chỉnh thông tin tác giả, đơn vị và nhấn nút **Bắt đầu định dạng** để tải xuống tệp kết quả hoàn chỉnh chuẩn mực.

### Cách 2: Sử dụng dòng lệnh (CLI Script)
Để định dạng nhanh chóng trực tiếp từ Terminal:
1.  Mở tệp `tool_dinh_dang/run_for_new_file.py`.
2.  Chỉnh sửa dữ liệu đầu vào (tên bài báo, tác giả, tóm tắt) ở `BƯỚC 1`.
3.  Chạy lệnh:
```bash
python tool_dinh_dang/run_for_new_file.py
```
Tệp kết quả sau định dạng sẽ được lưu tự động tại thư mục gốc dự án dưới dạng bản thảo hoàn chỉnh `v4.docx`.

---

## 🤝 Hỗ Trợ Đóng Góp Ý Kiến
Mọi ý kiến đóng góp, phát hiện lỗi hoặc yêu cầu thêm tính năng mới xin vui lòng tạo **Issue** hoặc gửi **Pull Request** trực tiếp trên kho lưu trữ GitHub của dự án!

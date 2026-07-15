import streamlit as st
import os
import io
import json
import tempfile
import google.generativeai as genai
import docx
from config import GEMINI_API_KEY
from extractor import extract_article_structure, extract_text_from_docx
from reviewer import run_ai_review
from formatter import format_document

# Thiết lập Page Config với nhận diện thương hiệu cao cấp
st.set_page_config(
    page_title="JSLHU Auto-Formatter & Reviewer",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Nhúng Custom CSS cao cấp theo phong cách Modern SaaS Dashboard
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #F8FAFC;
    }
    
    /* Thiết kế Header chứa Banner thương hiệu LHU */
    .header-container {
        background: linear-gradient(135deg, #15397F 0%, #2E5BFF 100%);
        padding: 2.5rem;
        border-radius: 20px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(21, 57, 127, 0.15);
        position: relative;
        overflow: hidden;
    }
    
    .header-container::after {
        content: "";
        position: absolute;
        top: -50%;
        right: -30%;
        width: 300px;
        height: 300px;
        background: rgba(255, 255, 255, 0.1);
        border-radius: 50%;
        filter: blur(60px);
    }
    
    .main-title {
        font-size: 2.6rem;
        font-weight: 700;
        margin: 0;
        color: white;
        letter-spacing: -0.5px;
    }
    
    .sub-title {
        font-size: 1.15rem;
        color: #E2E8F0;
        margin-top: 0.5rem;
        font-weight: 400;
    }
    
    /* Tùy chỉnh Sidebar phong cách tối (Dark mode) */
    [data-testid="stSidebar"] {
        background-color: #0F172A;
        color: white;
    }
    [data-testid="stSidebar"] h3 {
        color: #38BDF8 !important;
        font-weight: 600;
    }
    
    /* Thẻ thông tin cao cấp (Premium Card) */
    .premium-card {
        background: white;
        border-radius: 16px;
        padding: 1.5rem;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.02);
        margin-bottom: 1.5rem;
    }
    
    .status-card {
        background: #F1F5F9;
        border-left: 5px solid #15397F;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1.2rem;
    }
    
    /* Tùy chỉnh các Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
        border-bottom: 2px solid #E2E8F0;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: transparent;
        color: #64748B;
        font-size: 1.05rem;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    
    .stTabs [aria-selected="true"] {
        color: #15397F !important;
        border-bottom: 3px solid #15397F !important;
    }
    
    /* Huy hiệu chấm điểm tròn */
    .score-badge {
        font-size: 2.2rem;
        font-weight: 700;
        color: #15397F;
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        padding: 0.75rem 2.5rem;
        border-radius: 50px;
        display: inline-block;
        margin-bottom: 1.5rem;
        border: 1px solid #BFDBFE;
        box-shadow: 0 4px 12px rgba(21, 57, 127, 0.05);
    }
    
    /* Thẻ tiêu chí đánh giá */
    .criterion-item {
        padding: 1.1rem 1.4rem;
        border-radius: 12px;
        margin-bottom: 0.8rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.01);
        transition: all 0.2s ease;
    }
    
    .criterion-item:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.04);
    }
    
    .status-pass {
        background-color: #F0FDF4;
        color: #166534;
        border-left: 6px solid #22C55E;
    }
    
    .status-fail {
        background-color: #FEF2F2;
        color: #991B1B;
        border-left: 6px solid #EF4444;
    }
    
    .status-manual {
        background-color: #FFFBEB;
        color: #92400E;
        border-left: 6px solid #F59E0B;
    }
</style>
""", unsafe_allow_html=True)

# Hàm phân tích chuỗi Tác giả nhập từ form
def parse_authors_string(author_str):
    authors = []
    for part in author_str.split(","):
        part = part.strip()
        if not part:
            continue
        is_corr = "*" in part
        part = part.replace("*", "")
        
        # Tìm số mũ liên kết đơn vị công tác ở cuối tên
        name = part
        aff_num = ""
        for char in reversed(part):
            if char.isdigit():
                aff_num = char + aff_num
            else:
                break
        if aff_num:
            name = part[:-len(aff_num)].strip()
            
        authors.append({
            "name": name,
            "affiliation_number": aff_num,
            "is_corresponding": is_corr
        })
    return authors

# Khung tiêu đề chính của ứng dụng
st.markdown("""
<div class='header-container'>
    <div class='main-title'>📝 JSLHU Auto-Formatter & Reviewer</div>
    <div class='sub-title'>Hệ thống AI tự động hóa hoàn toàn việc trích xuất, định dạng bản thảo và chấm điểm 24 tiêu chí sơ loại - Tạp chí LHU</div>
</div>
""", unsafe_allow_html=True)

# Cấu hình thanh bên (Sidebar)
st.sidebar.markdown("### 🔑 Cấu hình hệ thống")
api_key = st.sidebar.text_input(
    "Gemini API Key",
    value=GEMINI_API_KEY,
    type="password",
    help="Nhập Gemini API Key của bạn để sử dụng AI phân tích và tự động định dạng bài báo."
)

st.sidebar.markdown("""
---
### 💡 Chế độ Tự Động Hóa 100%:
Không cần sửa bất kỳ dòng mã nào bằng tay!
1. Tải bản thảo thô của bạn lên ở **Tab 1**.
2. Nhấn nút **Trích xuất thông tin**.
3. **Kiểm tra và sửa trực tiếp** tên tác giả/đơn vị công tác trên biểu mẫu nếu file thô bị thiếu.
4. Nhấn **Định dạng & Chấm điểm** để tải tệp Word chuẩn hai cột về!
""")

# Các Tab chính
tab1, tab2 = st.tabs(["🔍 Định dạng & Đánh giá Bản thảo tự động (Khuyên dùng)", "🚀 Viết bài báo từ Đề cương (Outline)"])

# ----------------- TAB 1: ĐỊNH DẠNG & ĐÁNH GIÁ BẢN THẢO TỰ ĐỘNG -----------------
with tab1:
    st.subheader("📋 Tải lên và tự động hóa định dạng bản thảo thô")
    st.write(
        "Chỉ cần tải tệp Word thô của bạn lên (ví dụ: `ungdung.docx`). AI sẽ tự động đọc, trích xuất "
        "thông tin ban đầu, sau đó cho phép bạn kiểm tra/sửa lại trước khi ráp vào tệp mẫu của LHU."
    )
    
    # Khởi tạo trạng thái phiên làm việc cho Tab 1
    if "processed_data" not in st.session_state:
        st.session_state.processed_data = None
    if "review_results" not in st.session_state:
        st.session_state.review_results = None
    if "output_docx_bytes" not in st.session_state:
        st.session_state.output_docx_bytes = None
    if "input_file_name" not in st.session_state:
        st.session_state.input_file_name = None
    if "input_file_bytes" not in st.session_state:
        st.session_state.input_file_bytes = None

    col1, col2 = st.columns([1.1, 0.9])
    
    with col1:
        st.markdown("<div class='premium-card'>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Tải lên tệp bản thảo bài báo (.docx)",
            type=["docx"],
            key="draft_uploader",
            help="Hãy chọn file bản thảo thô của bạn để tiến hành định dạng."
        )
        st.markdown("</div>", unsafe_allow_html=True)
        
        if uploaded_file is not None:
            # Lưu trữ file bytes vào session state để dùng cho các bước sau
            if st.session_state.input_file_name != uploaded_file.name:
                st.session_state.processed_data = None
                st.session_state.review_results = None
                st.session_state.output_docx_bytes = None
                st.session_state.input_file_name = uploaded_file.name
                st.session_state.input_file_bytes = uploaded_file.read()
            
            st.markdown(f"""
            <div class='status-card'>
                <h4 style='margin:0 0 0.5rem 0; color:#15397F;'>📁 Thông tin file bản thảo</h4>
                <p style='margin:0.2rem 0;'><b>Tên tệp:</b> {st.session_state.input_file_name}</p>
                <p style='margin:0.2rem 0;'><b>Dung lượng:</b> {round(len(st.session_state.input_file_bytes) / 1024, 2)} KB</p>
            </div>
            """, unsafe_allow_html=True)
            
            # BƯỚC 1: Nút trích xuất thông tin bằng AI
            if st.session_state.processed_data is None:
                if st.button("🔍 Bước 1: Trích xuất thông tin bài báo bằng AI", use_container_width=True):
                    if not api_key:
                        st.error("⚠️ Vui lòng cung cấp Gemini API Key ở thanh công cụ bên trái trước khi bắt đầu!")
                    else:
                        with st.spinner("Đang sử dụng AI (Gemini) để đọc và tự động bóc tách cấu trúc bài viết..."):
                            data = extract_article_structure(io.BytesIO(st.session_state.input_file_bytes), api_key)
                            st.session_state.processed_data = data
                            st.rerun()
            
            # BƯỚC 2: Biểu mẫu điền/sửa thông tin nếu AI đã trích xuất xong
            if st.session_state.processed_data is not None:
                st.markdown("<div class='premium-card'>", unsafe_allow_html=True)
                st.subheader("✏️ Chỉnh sửa thông tin bài báo (Metadata)")
                st.info("💡 Bạn có thể điền hoặc điều chỉnh lại thông tin dưới đây (nhất là nếu tệp thô ban đầu bị thiếu tác giả, đơn vị công tác hoặc email).")
                
                ext_data = st.session_state.processed_data
                
                # Tiêu đề
                title_vi = st.text_input("Tiêu đề tiếng Việt (Viết HOA)", value=ext_data.get("title_vi", ""))
                title_en = st.text_input("Tiêu đề tiếng Anh (Viết HOA)", value=ext_data.get("title_en", ""))
                
                # Tác giả tiếng Việt / tiếng Anh
                authors_vi_list = ext_data.get("authors_vi", [])
                if not authors_vi_list or any("Tác giả" in a.get("name", "") or "Tác gi" in a.get("name", "") for a in authors_vi_list):
                    default_authors_vi = "Bùi Xuân Cảnh1*, Nguyễn Hoàng Anh1"
                else:
                    default_authors_vi = ", ".join([f"{a.get('name', '')}{a.get('affiliation_number', '')}{'*' if a.get('is_corresponding') else ''}" for a in authors_vi_list])
                    
                authors_en_list = ext_data.get("authors_en", [])
                if not authors_en_list or any("Author" in a.get("name", "") for a in authors_en_list):
                    default_authors_en = "Bui Xuan Canh1*, Nguyen Hoang Anh1"
                else:
                    default_authors_en = ", ".join([f"{a.get('name', '')}{a.get('affiliation_number', '')}{'*' if a.get('is_corresponding') else ''}" for a in authors_en_list])
                
                authors_vi_str = st.text_input("Danh sách tác giả tiếng Việt (Ví dụ: Bùi Xuân Cảnh1*, Nguyễn Hoàng Anh1)", value=default_authors_vi)
                authors_en_str = st.text_input("Danh sách tác giả tiếng Anh (Ví dụ: Bui Xuan Canh1*, Nguyen Hoang Anh1)", value=default_authors_en)
                
                # Đơn vị công tác tiếng Việt / tiếng Anh
                affiliations_vi_list = ext_data.get("affiliations_vi", [])
                if not affiliations_vi_list:
                    default_aff_vi = "Trường Đại học Lạc Hồng, Số 10 Đường Huỳnh văn Nghệ, Phường Trấn Biên, Thành phố Đồng Nai, Việt Nam"
                else:
                    default_aff_vi = "\n".join([a.get("name", "") for a in affiliations_vi_list])
                    
                affiliations_en_list = ext_data.get("affiliations_en", [])
                if not affiliations_en_list:
                    default_aff_en = "Lac Hong University, No. 10 Huynh Van Nghe Street, Tran Bien Ward, Dong Nai City, Vietnam"
                else:
                    default_aff_en = "\n".join([a.get("name", "") for a in affiliations_en_list])
                
                aff_vi_text = st.text_area("Đơn vị công tác tiếng Việt (Mỗi đơn vị viết một dòng)", value=default_aff_vi, height=80)
                aff_en_text = st.text_area("Đơn vị công tác tiếng Anh (Mỗi đơn vị viết một dòng)", value=default_aff_en, height=80)
                
                # Email liên hệ
                email_contact_val = ext_data.get("email_contact", "")
                if not email_contact_val or "nguyenvana" in email_contact_val.lower():
                    email_contact_val = "Canhbx@lhu.edu.vn"
                
                email_contact = st.text_input("Email tác giả liên hệ", value=email_contact_val)
                
                # Tóm tắt
                abstract_vi = st.text_area("Nội dung tóm tắt tiếng Việt", value=ext_data.get("abstract_vi", ""), height=150)
                abstract_en = st.text_area("Nội dung tóm tắt tiếng Anh (ABSTRACT)", value=ext_data.get("abstract_en", ""), height=150)
                
                # Từ khóa
                default_kw_vi = ", ".join(ext_data.get("keywords_vi", []))
                default_kw_en = ", ".join(ext_data.get("keywords_en", []))
                keywords_vi_str = st.text_input("Từ khóa tiếng Việt (phân cách bằng dấu phẩy)", value=default_kw_vi)
                keywords_en_str = st.text_input("Từ khóa tiếng Anh (phân cách bằng dấu phẩy)", value=default_kw_en)
                
                st.markdown("</div>", unsafe_allow_html=True)
                
                # Nút chạy xuất file Word và Chấm điểm dựa trên thông tin vừa sửa
                if st.button("🚀 Bước 2: Định dạng bản thảo & Đánh giá sơ loại", use_container_width=True):
                    try:
                        # Chuẩn hóa lại cấu trúc dữ liệu gửi vào formatter
                        parsed_authors_vi = parse_authors_string(authors_vi_str)
                        parsed_authors_en = parse_authors_string(authors_en_str)
                        
                        parsed_aff_vi = [{"number": str(i+1), "name": line.strip()} for i, line in enumerate(aff_vi_text.split("\n")) if line.strip()]
                        parsed_aff_en = [{"number": str(i+1), "name": line.strip()} for i, line in enumerate(aff_en_text.split("\n")) if line.strip()]
                        
                        parsed_kw_vi = [k.strip() for k in keywords_vi_str.split(",") if k.strip()]
                        parsed_kw_en = [k.strip() for k in keywords_en_str.split(",") if k.strip()]
                        
                        formatted_data = {
                            "title_vi": title_vi,
                            "title_en": title_en,
                            "authors_vi": parsed_authors_vi,
                            "authors_en": parsed_authors_en,
                            "affiliations_vi": parsed_aff_vi,
                            "affiliations_en": parsed_aff_en,
                            "email_contact": email_contact,
                            "abstract_vi": abstract_vi,
                            "abstract_en": abstract_en,
                            "keywords_vi": parsed_kw_vi,
                            "keywords_en": parsed_kw_en,
                            "sections": ext_data.get("sections", []),
                            "acknowledgments": ext_data.get("acknowledgments", ""),
                            "references": ext_data.get("references", [])
                        }
                        
                        with st.spinner("Đang tiến hành chấm điểm sơ loại 24 tiêu chí LHU..."):
                            review = run_ai_review(formatted_data, api_key)
                            st.session_state.review_results = review
                            
                        with st.spinner("Đang tạo tệp Word mới, phân cột và tối ưu hóa độ rộng cột..."):
                            template_path = "BM01. jslhu template.docx"
                            if not os.path.exists(template_path):
                                st.error(f"Không tìm thấy file mẫu template '{template_path}' trong thư mục dự án!")
                            else:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                                    tmp_output_path = tmp_file.name
                                
                                # Tạo file tạm của bản gốc để formatter có thể trích xuất ảnh/bảng
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_src:
                                    tmp_src.write(st.session_state.input_file_bytes)
                                    tmp_src_path = tmp_src.name
                                
                                try:
                                    format_document(formatted_data, template_path, tmp_output_path, source_docx_path=tmp_src_path)
                                    
                                    with open(tmp_output_path, "rb") as f:
                                        st.session_state.output_docx_bytes = f.read()
                                finally:
                                    if os.path.exists(tmp_output_path):
                                        os.remove(tmp_output_path)
                                    if os.path.exists(tmp_src_path):
                                        os.remove(tmp_src_path)
                                
                                st.success("🎉 Xử lý định dạng và chấm điểm bài viết thành công!")
                                st.rerun()
                                
                    except Exception as e:
                        st.error(f"Đã xảy ra lỗi trong quá trình định dạng: {e}")
            
            # Hiển thị nút tải xuống nếu đã định dạng xong
            if st.session_state.output_docx_bytes is not None:
                st.markdown("### 📥 Kết quả đầu ra")
                st.download_button(
                    label="📥 Tải xuống bài báo đã định dạng chuẩn (BM01) (.docx)",
                    data=st.session_state.output_docx_bytes,
                    file_name=f"BM01_formatted_{st.session_state.input_file_name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.info("💡 Vui lòng tải lên file bản thảo của bạn dưới dạng Word (.docx) ở trên để bắt đầu.")
            
    with col2:
        st.subheader("📊 Bảng đánh giá sơ loại bài báo")
        
        if st.session_state.review_results is not None:
            review = st.session_state.review_results
            results = review.get("results", [])
            score = review.get("score", 0)
            max_score = review.get("max_score", 24)
            
            st.markdown(f"""
            <div style='text-align: center;'>
                <div class='score-badge'>Đạt {score} / {max_score} tiêu chí</div>
            </div>
            """, unsafe_allow_html=True)
            
            pass_count = sum(1 for item in results if item.get("status") is True)
            fail_count = sum(1 for item in results if item.get("status") is False)
            manual_count = sum(1 for item in results if item.get("status") is None)
            
            col_p, col_f, col_m = st.columns(3)
            col_p.metric("Đạt", f"{pass_count} / 24")
            col_f.metric("Chưa đạt", f"{fail_count} / 24")
            col_m.metric("Cần kiểm tra", f"{manual_count} / 24")
            
            st.markdown("### Chi tiết kết quả chấm 24 tiêu chí:")
            
            for item in results:
                crit_id = item.get("id", "")
                crit_name = item.get("criterion", "")
                status = item.get("status")
                note = item.get("note", "")
                
                if status is True:
                    status_class = "status-pass"
                    status_text = "✓ Đạt"
                elif status is False:
                    status_class = "status-fail"
                    status_text = "✗ Chưa đạt"
                else:
                    status_class = "status-manual"
                    status_text = "⚠ Kiểm tra"
                    
                st.markdown(f"""
                <div class='criterion-item {status_class}'>
                    <div>
                        <b>{crit_id}. {crit_name}</b><br>
                        <small style='color: #5F6368;'>Ghi chú: {note}</small>
                    </div>
                    <div style='font-weight: bold;'>{status_text}</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("📊 Kết quả chấm điểm sơ loại bài viết của bạn sẽ hiển thị ở đây sau khi chạy xử lý.")

# ----------------- TAB 2: SINH BÀI BÁO TỪ ĐỀ CƯƠNG -----------------
with tab2:
    st.subheader("💡 Tự động viết & định dạng bài báo từ Đề cương chi tiết (Outline)")
    st.write(
        "Chức năng này cho phép bạn tải lên đề cương thô (Outline). AI sẽ tự động phân tích cấu trúc đề cương, "
        "mở rộng thành bài viết khoa học chi tiết có chiều sâu kỹ thuật, số liệu thực nghiệm nhất quán, "
        "và xuất thẳng ra mẫu file Word `BM01` chuẩn hai cột."
    )
    
    outline_file = st.file_uploader(
        "Tải lên tệp đề cương Word (.docx)",
        type=["docx"],
        key="outline_uploader",
        help="Hãy chọn file đề cương bài viết của bạn."
    )
    
    outline_text_input = st.text_area(
        "Hoặc dán nội dung đề cương tại đây (nếu không có file):",
        height=300,
        placeholder="Dán nội dung đề cương chi tiết tại đây..."
    )
    
    if st.button("🚀 Bắt đầu sinh bài báo hoàn chỉnh từ Outline", use_container_width=True):
        if not api_key:
            st.error("⚠️ Vui lòng cung cấp Gemini API Key ở thanh công cụ bên trái trước khi bắt đầu!")
        else:
            try:
                outline_content = ""
                if outline_file is not None:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        tmp.write(outline_file.read())
                        tmp_path = tmp.name
                    outline_content = extract_text_from_docx(tmp_path)
                    os.remove(tmp_path)
                elif outline_text_input.strip():
                    outline_content = outline_text_input.strip()
                    
                if not outline_content:
                    st.error("⚠️ Vui lòng tải lên file đề cương hoặc dán nội dung vào ô văn bản!")
                else:
                    with st.spinner("Đang kết nối Gemini API để lập luận và viết bài báo chi tiết... (Quá trình này có thể mất 1-2 phút)"):
                        genai.configure(api_key=api_key)
                        
                        prompt = f"""
                        Bạn là một chuyên gia nghiên cứu khoa học và biên tập viên tạp chí cấp cao.
                        Nhiệm vụ của bạn là lấy đề cương chi tiết (Outline) dưới đây và mở rộng nó thành một bài báo nghiên cứu khoa học đầy đủ, chi tiết, đi sâu phân tích kỹ thuật công nghệ và có số liệu thực nghiệm thực tế chặt chẽ bằng tiếng Việt.
                        
                        YÊU CẦU NỘI DUNG CHI TIẾT:
                        - Viết các đoạn văn nghị luận dài, sâu sắc cho từng chương mục chính (Giới thiệu, Phương pháp nghiên cứu, Kết quả thực nghiệm và Thảo luận, Kết luận).
                        - Không viết tóm tắt ngắn gọn. Mỗi mục cần có lập luận khoa học, dẫn chứng, giải thích rõ cơ chế vận hành của công nghệ (Ví dụ: RAG, gemini-embedding-001, vector 768 chiều, chỉ mục HNSW, PGVector, xoay vòng API Keys, bộ lọc tương đồng Min-Max delta = 0.65, kiểm định độ tin cậy Cronbach's Alpha và Fleiss' Kappa đồng thuận giáo viên).
                        
                        Đầu ra của bạn bắt buộc phải là định dạng JSON thuần túy (không có thẻ ```json ở đầu hoặc cuối) theo đúng cấu trúc sau:
                        {{
                          "title_vi": "TIÊU ĐỀ TIẾNG VIỆT IN HOA",
                          "title_en": "TITLE IN ENGLISH IN CAPITAL",
                          "authors_vi": [
                             {{"name": "Bùi Xuân Cảnh", "affiliation_number": "1", "is_corresponding": true}},
                             {{"name": "Nguyễn Hoàng Anh", "affiliation_number": "1", "is_corresponding": false}}
                          ],
                          "authors_en": [
                             {{"name": "Bui Xuan Canh", "affiliation_number": "1", "is_corresponding": true}},
                             {{"name": "Nguyen Hoang Anh", "affiliation_number": "1", "is_corresponding": false}}
                          ],
                          "affiliations_vi": [
                             {{"number": "1", "name": "Trường Đại học Lạc Hồng, Số 10 Đường Huỳnh văn Nghệ, Phường Trấn Biên, Thành phố Đồng Nai, Việt Nam"}}
                          ],
                          "affiliations_en": [
                             {{"number": "1", "name": "Lac Hong University, No. 10 Huynh Van Nghe Street, Tran Bien Ward, Dong Nai City, Vietnam"}}
                          ],
                          "email_contact": "Canhbx@lhu.edu.vn",
                          "abstract_vi": "Đoạn văn tóm tắt bằng tiếng Việt (150-250 từ).",
                          "abstract_en": "ABSTRACT in English (150-250 words).",
                          "keywords_vi": ["RAG thích ứng", "Gia sư ảo", "Supabase", "PGVector", "Cá nhân hóa"],
                          "keywords_en": ["Adaptive RAG", "Virtual Tutor", "Supabase", "PGVector", "Personalization"],
                          "sections": [
                             {{
                               "title": "1. Giới thiệu",
                               "content": "Nội dung phần giới thiệu viết cực kỳ dài và học thuật..."
                             }},
                             {{
                               "title": "2. Phương pháp nghiên cứu",
                               "content": "Nội dung phần phương pháp đi sâu vào các công nghệ và giải thuật kỹ thuật..."
                             }},
                             {{
                               "title": "3. Kết quả thực nghiệm và Thảo luận",
                               "content": "Nội dung phần kết quả thực nghiệm, số liệu khảo sát, và thảo luận..."
                             }},
                             {{
                               "title": "4. Kết luận",
                               "content": "Nội dung kết luận và đề xuất hướng Graph-RAG..."
                             }}
                          ],
                          "acknowledgments": "Lời cảm ơn (nếu có)",
                          "references": [
                             "[1] Phan Van Nam, Nguyen Trung Hieu. Mitigating disparities in legal literacy: A RAG-based system for navigating the Vietnamese legal framework. Journal of Science of Lac Hong University, 2026, 26, 076-081.",
                             "[2] Gao Y., et al. Retrieval-augmented generation for large language models: A survey. arXiv preprint arXiv:2312.10997, 2023."
                          ]
                        }}
                        
                        NỘI DUNG ĐỀ CƯƠNG CHI TIẾT (OUTLINE) ĐẦU VÀO:
                        {outline_content}
                        """
                        
                        model = genai.GenerativeModel("gemini-2.5-flash")
                        response = model.generate_content(
                            prompt,
                            generation_config={"response_mime_type": "application/json"}
                        )
                        
                        clean_text = response.text.strip()
                        if clean_text.startswith("```json"):
                            clean_text = clean_text[7:]
                        if clean_text.endswith("```"):
                            clean_text = clean_text[:-3]
                        
                        generated_data = json.loads(clean_text.strip())
                        
                    with st.spinner("Đang tiến hành tự động định dạng bài viết vừa viết thành file Word BM01..."):
                        template_path = "BM01. jslhu template.docx"
                        if not os.path.exists(template_path):
                            st.error(f"Không tìm thấy file mẫu template '{template_path}' trong thư mục dự án!")
                        else:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                                tmp_output_path = tmp_file.name
                            
                            format_document(generated_data, template_path, tmp_output_path)
                            
                            with open(tmp_output_path, "rb") as f:
                                doc_bytes = f.read()
                            os.remove(tmp_output_path)
                            
                            st.success("🎉 Tự động viết và định dạng bài báo thành công!")
                            st.download_button(
                                label="📥 Tải xuống bài báo hoàn chỉnh định dạng BM01 (.docx)",
                                data=doc_bytes,
                                file_name="BM01_Gia_Su_Ao_RAG_K12.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
            except Exception as e:
                st.error(f"Đã xảy ra lỗi trong quá trình tự động sinh bài viết: {e}")

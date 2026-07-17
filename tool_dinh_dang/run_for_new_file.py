import sys
import os
from formatter import format_document

# =====================================================================
# BƯỚC 1: ĐIỀN THÔNG TIN BÀI BÁO MỚI CỦA BẠN VÀO ĐÂY
# =====================================================================
data = {
    # Tiêu đề bài viết (Viết HOA)
    "title_vi": "NGHIÊN CỨU KIẾN TRÚC RAG THÍCH ỨNG TRONG TỐI ƯU HÓA PHƯƠNG PHÁP DẠY HỌC CÁ NHÂN HÓA K-12",
    "title_en": "RESEARCH ON ADAPTIVE RAG ARCHITECTURE FOR OPTIMIZING PERSONALIZED K-12 TEACHING METHODOLOGY",
    
    # Danh sách tác giả tiếng Việt (Bùi Xuân Cảnh, Nguyễn Hoàng Anh và Trần Hoài Phương mang số mũ 1, Cảnh là tác giả liên hệ)
    "authors_vi": [
        {"name": "Bùi Xuân Cảnh", "affiliation_number": "1", "is_corresponding": True},
        {"name": "Nguyễn Hoàng Anh", "affiliation_number": "1", "is_corresponding": False},
        {"name": "Trần Hoài Phương", "affiliation_number": "1", "is_corresponding": False}
    ],
    
    # Danh sách tác giả tiếng Anh tương ứng
    "authors_en": [
        {"name": "Bui Xuan Canh", "affiliation_number": "1", "is_corresponding": True},
        {"name": "Nguyen Hoang Anh", "affiliation_number": "1", "is_corresponding": False},
        {"name": "Tran Hoai Phuong", "affiliation_number": "1", "is_corresponding": False}
    ],
    
    # Danh sách đơn vị công tác tiếng Việt
    "affiliations_vi": [
        {"number": "1", "name": "Trường Đại học Lạc Hồng, Số 10, Huỳnh Văn Nghệ, phường Trấn Biên, thành phố Đồng Nai, Việt Nam"}
    ],
    
    # Danh sách đơn vị công tác tiếng Anh
    "affiliations_en": [
        {"number": "1", "name": "Lac Hong University, No. 10, Huynh Van Nghe Street, Tran Bien Ward, Dong Nai City, Vietnam"}
    ],
    
    # Email liên hệ của tác giả chính
    "email_contact": "Canhbx@lhu.edu.vn",
    
    # Tóm tắt tiếng Việt (Nhập trực tiếp nội dung, không ghi chữ "TÓM TẮT")
    "abstract_vi": "Bài báo này nghiên cứu giải quyết bất bình đẳng tiếp cận học liệu số ở học sinh phổ thông từ lớp một đến lớp mười hai bằng cách đề xuất kiến trúc truy xuất tăng cường tạo sinh thích ứng phân cấp để tối ưu hóa dạy học cá nhân hóa. Phương pháp nghiên cứu bao gồm thiết kế bộ định tuyến câu hỏi để phân lớp người học, kết hợp tìm kiếm lai giữa tần suất từ khóa và độ tương đồng ngữ nghĩa, tái xếp hạng bằng mô hình học sâu và điều hướng phản hồi theo phương pháp gợi mở từng bước kết hợp hồ sơ năng lực động của học sinh theo vùng phát triển gần nhất. Kết quả thực nghiệm đối chứng trên sáu mươi học sinh giúp tăng một trăm năm mươi mốt phẩy một phần trăm thời gian tự học tập trung, đồng thời giảm tỷ lệ sao chép lời giải xuống mười phẩy năm phần trăm. Đánh giá định lượng qua bộ chỉ số tiêu chuẩn khẳng định hệ thống đạt độ trung thực chín mươi tư phần trăm và độ phù hợp chín mươi mốt phần trăm. Chỉ số đồng thuận của mười giáo viên đánh giá đạt không phẩy bảy mươi sáu. Nghiên cứu này chứng minh tính khả thi của trợ lý học thuật trí tuệ nhân tạo trong việc nâng cao năng lực tự học cá nhân hóa của học sinh.",
    
    # Tóm tắt tiếng Anh (Nhập trực tiếp nội dung, không ghi chữ "ABSTRACT")
    "abstract_en": "This paper investigates addressing inequality in access to digital learning materials for school students from grade one to grade twelve by proposing a hierarchical adaptive retrieval-augmented generation architecture to optimize personalized learning. The research methodology includes designing a query router to classify learners, combining hybrid search between term frequency and semantic similarity, reranking using a deep learning model, and guiding feedback through a step-by-step scaffolding method combined with students' dynamic competency profiles according to the zone of proximal development. Experimental results on sixty students show an increase of one hundred fifty-one point one percent in focused self-study time, while reducing solution copying to ten point five percent. Quantitative evaluation confirms that the system achieves ninety-four percent faithfulness and ninety-one percent answer relevance. The consensus index of ten evaluating teachers reaches zero point seventy-six. This study demonstrates the feasibility of artificial intelligence academic assistants in enhancing students' personalized self-study capabilities.",
    
    # Từ khóa tiếng Việt (Dạng mảng các từ khóa)
    "keywords_vi": ["RAG thích ứng", "Gia sư ảo K-12", "PGVector", "Lọc Metadata", "Sư phạm gợi mở"],
    
    # Từ khóa tiếng Anh tương ứng
    "keywords_en": ["Adaptive RAG", "K-12 Virtual Tutor", "PGVector", "Metadata Filtering", "Scaffolding Pedagogy"]
}

# =====================================================================
# BƯỚC 2: THAY ĐỔI ĐƯỜNG DẪN TỆP CỦA BẠN TẠI ĐÂY
# =====================================================================
# Tệp nguồn bản thảo thô mới của bạn (Nội dung chỉ cần bắt đầu từ phần "1. Giới thiệu" hoặc "Introduction")
base_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(base_dir)

FILE_NGUON_THO = os.path.join(project_root, "bao_cao_nghien_cuu_RAG_hoan_chinh_tieng_viet.docx")

# Đường dẫn muốn lưu tệp kết quả sau định dạng
FILE_KET_QUA = os.path.join(project_root, "bao_cao_nghien_cuu_RAG_hoan_chinh_v4.docx")

# Đường dẫn tệp template (Giữ nguyên tệp mẫu này)
FILE_TEMPLATE = os.path.join(project_root, "BM01. jslhu template.docx")


# =====================================================================
# BƯỚC 3: CHẠY ĐỊNH DẠNG
# =====================================================================
if __name__ == "__main__":
    if not os.path.exists(FILE_NGUON_THO):
        print(f"ERROR: Raw source file not found at: {FILE_NGUON_THO}")
        print("Please check the path FILE_NGUON_THO in Step 2.")
    else:
        print("--------------------------------------------------")
        print("Processing document formatting...")
        try:
            format_document(data, FILE_TEMPLATE, FILE_KET_QUA, source_docx_path=FILE_NGUON_THO)
            print("--------------------------------------------------")
            print("FORMATTING SUCCESSFUL!")
            print("Output saved to:")
            print(f"-> {FILE_KET_QUA.encode('utf-8')}")
        except Exception as e:
            import traceback
            traceback.print_exc()

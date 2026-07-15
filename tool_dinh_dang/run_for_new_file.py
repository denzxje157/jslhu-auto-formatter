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
    
    # Danh sách tác giả tiếng Việt (Bùi Xuân Cảnh và Nguyễn Hoàng Anh mang số mũ 1, Cảnh là tác giả liên hệ)
    "authors_vi": [
        {"name": "Bùi Xuân Cảnh", "affiliation_number": "1", "is_corresponding": True},
        {"name": "Nguyễn Hoàng Anh", "affiliation_number": "1", "is_corresponding": False}
    ],
    
    # Danh sách tác giả tiếng Anh tương ứng
    "authors_en": [
        {"name": "Bui Xuan Canh", "affiliation_number": "1", "is_corresponding": True},
        {"name": "Nguyen Hoang Anh", "affiliation_number": "1", "is_corresponding": False}
    ],
    
    # Danh sách đơn vị công tác tiếng Việt
    "affiliations_vi": [
        {"number": "1", "name": "Trường Đại học Lạc Hồng, Số 10 đường Huỳnh Văn Nghệ, phường Bửu Long, thành phố Biên Hòa, tỉnh Đồng Nai, Việt Nam"}
    ],
    
    # Danh sách đơn vị công tác tiếng Anh
    "affiliations_en": [
        {"number": "1", "name": "Lac Hong University, No. 10, Huynh Van Nghe Street, Buu Long Ward, Bien Hoa City, Dong Nai Province, Vietnam"}
    ],
    
    # Email liên hệ của tác giả chính
    "email_contact": "Canhbx@lhu.edu.vn",
    
    # Tóm tắt tiếng Việt (Nhập trực tiếp nội dung, không ghi chữ "TÓM TẮT")
    "abstract_vi": "Bài báo này nghiên cứu giải quyết sự bất bình đẳng tiếp cận học liệu số ở học sinh phổ thông (K-12) bằng cách đề xuất kiến trúc RAG thích ứng phân cấp. Để khắc phục lỗi nhận diện thuật ngữ và ảo giác kiến thức, chúng tôi tích hợp cơ chế tìm kiếm kết hợp BM25 và Cosine, tối ưu hóa bằng thuật toán xếp hạng hỗn hợp RRF và tái xếp hạng Cross-Encoder BGE-Reranker-Large. Hệ thống sử dụng cơ sở dữ liệu Supabase PGVector và cấu hình prompt Socratic để định hướng tư duy học sinh theo vùng phát triển gần nhất. Kết quả thực nghiệm đối chứng trên 60 học sinh cho thấy hệ thống tăng 151,1% thời gian tự học hiệu quả, giảm tỷ lệ sao chép lời giải xuống 10,5%. Đánh giá định lượng Ragas khẳng định hệ thống đạt độ trung thực 94% và độ phù hợp 91%. Nghiên cứu chứng minh tính khả thi của trợ lý AI học thuật trong giáo dục phổ thông cá nhân hóa.",
    
    # Tóm tắt tiếng Anh (Nhập trực tiếp nội dung, không ghi chữ "ABSTRACT")
    "abstract_en": "This paper addresses the educational inequality in K-12 learning by proposing a hierarchical adaptive RAG architecture. To resolve equation parsing errors and knowledge hallucination, we integrate a hybrid search mechanism combining BM25 and Cosine similarity, optimized by Reciprocal Rank Fusion (RRF) and Cross-Encoder BGE-Reranker-Large. The system utilizes a Supabase PGVector database and a Socratic prompting hierarchy to guide student reasoning within the Zone of Proximal Development. Experimental results on 60 students demonstrate a 151.1% increase in self-study time and a reduction in solution-copying habits to 10.5%. Quantitative evaluation using Ragas confirms that the system achieves 94% faithfulness and 91% answer relevance. This research demonstrates the feasibility of academic AI tutors in personalized K-12 education.",
    
    # Từ khóa tiếng Việt (Dạng mảng các từ khóa)
    "keywords_vi": ["RAG thích ứng", "Gia sư ảo K-12", "Supabase", "PGVector", "Lọc Metadata", "Sư phạm gợi mở"],
    
    # Từ khóa tiếng Anh tương ứng
    "keywords_en": ["Adaptive RAG", "K-12 Virtual Tutor", "Supabase", "PGVector", "Metadata Filtering", "Scaffolding Pedagogy"]
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

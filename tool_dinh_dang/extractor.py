import google.generativeai as genai
import json
import docx
import io

def extract_text_from_docx(file_stream_or_path):
    """
    Trích xuất toàn bộ text thô từ file docx bao gồm cả paragraphs và tables.
    """
    doc = docx.Document(file_stream_or_path)
    paragraphs = []
    
    # Duyệt qua các phần tử của tài liệu để giữ nguyên thứ tự tương đối
    for block in doc.element.body:
        if block.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(block, doc)
            if p.text.strip():
                paragraphs.append(p.text)
        elif block.tag.endswith('tbl'):
            t = docx.table.Table(block, doc)
            table_text = []
            for row in t.rows:
                row_vals = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                # Loại bỏ các cột trống trùng lặp do merged cells
                unique_vals = []
                for val in row_vals:
                    if not unique_vals or val != unique_vals[-1]:
                        unique_vals.append(val)
                table_text.append(" | ".join(unique_vals))
            paragraphs.append("\n[BẢNG THÔNG TIN/DỮ LIỆU]:\n" + "\n".join(table_text) + "\n[HẾT BẢNG]")
            
    return "\n\n".join(paragraphs)

def extract_article_structure(docx_file, api_key):
    """
    Gửi văn bản bài báo thô lên Gemini để trích xuất cấu trúc JSON chuẩn hóa.
    """
    genai.configure(api_key=api_key)
    
    # Đọc văn bản thô từ file
    raw_text = extract_text_from_docx(docx_file)
    
    prompt = f"""
    Bạn là một trợ lý khoa học chuyên nghiệp chuyên phân tích cấu trúc bài báo.
    Nhiệm vụ của bạn là phân tích toàn bộ văn bản của bài báo thô được cung cấp dưới đây, sau đó bóc tách và phân loại nó thành một cấu trúc JSON chính xác theo đặc tả dưới đây.

    Hãy đảm bảo dịch chuyển ngữ chính xác thông tin tiếng Anh và tiếng Việt tương ứng. Nếu bài viết thiếu phần tiếng Anh (ví dụ thiếu Title tiếng Anh hay Abstract tiếng Anh), bạn hãy tự chuyển ngữ từ tiếng Việt sang tiếng Anh cho phần đó để điền đầy đủ vào cấu trúc.

    Cấu trúc JSON yêu cầu trả về:
    {{
      "title_vi": "Tiêu đề bài báo bằng tiếng Việt (in hoa, không quá 20 từ)",
      "title_en": "Tiêu đề bài báo bằng tiếng Anh (in hoa, dịch từ tiếng Việt nếu không có sẵn)",
      "authors_vi": [
         {{"name": "Họ và tên tác giả 1", "affiliation_number": "1", "is_corresponding": true}},
         {{"name": "Họ và tên tác giả 2", "affiliation_number": "2", "is_corresponding": false}}
      ],
      "authors_en": [
         {{"name": "First Author Name", "affiliation_number": "1", "is_corresponding": true}},
         {{"name": "Second Author Name", "affiliation_number": "2", "is_corresponding": false}}
      ],
      "affiliations_vi": [
         {{"number": "1", "name": "Đơn vị công tác tiếng Việt thứ 1 (ví dụ: Khoa Công nghệ thông tin, Trường Đại học Lạc Hồng)"}},
         {{"number": "2", "name": "Đơn vị công tác tiếng Việt thứ 2"}}
      ],
      "affiliations_en": [
         {{"number": "1", "name": "Đơn vị công tác tiếng Anh thứ 1"}},
         {{"number": "2", "name": "Đơn vị công tác tiếng Anh thứ 2"}}
      ],
      "email_contact": "email_tác_giả_liên_hệ (nếu là tác giả thuộc LHU thì ưu tiên sử dụng email @lhu.edu.vn, tìm kiếm trong bài báo thô)",
      "abstract_vi": "Đoạn văn tóm tắt bằng tiếng Việt (khoảng 150-250 từ, mô tả mục tiêu, phương pháp, kết quả, kết luận. Không viết tắt, không trích dẫn tài liệu tham khảo)",
      "abstract_en": "Đoạn văn tóm tắt bằng tiếng Anh (ABSTRACT, khoảng 150-250 từ, dịch từ bản tiếng Việt nếu bài thô không có sẵn)",
      "keywords_vi": ["từ khóa 1", "từ khóa 2", "từ khóa 3", "từ khóa 4", "từ khóa 5"],
      "keywords_en": ["keyword 1", "keyword 2", "keyword 3", "keyword 4", "keyword 5"],
      "sections": [
         {{
           "title": "Tiêu đề chương mục chính (ví dụ: '1. Giới thiệu', '2. Phương pháp nghiên cứu', '3. Kết quả và Thảo luận', '4. Kết luận')",
           "content": "Toàn bộ nội dung văn bản của chương mục này. Nếu chứa nhiều đoạn văn thì ngăn cách nhau bằng dấu xuống dòng \\n"
         }}
      ],
      "acknowledgments": "Lời cảm ơn (để trống nếu bài báo không có)",
      "references": [
         "Dòng tài liệu tham khảo thứ 1, ví dụ: [1] Cranford S-W, Buehler M-J. Mechanical Properties of Graphyne. 2011, 49 (1), pp. 4111-4121.",
         "Dòng tài liệu tham khảo thứ 2..."
      ]
    }}

    YÊU CẦU NGHIÊM NGẶT:
    1. Trả về định dạng JSON thuần túy, không có thẻ ```json ở đầu hoặc cuối, không có bất kỳ lời giải thích nào ngoài JSON.
    2. Cố gắng trích xuất đầy đủ và chính xác tất cả thông tin. Đối với phần 'sections', hãy thu thập toàn bộ văn bản của bài báo từ đầu đến trước phần Tài liệu tham khảo, nhóm chúng theo các đề mục lớn tương ứng.
    3. Trích xuất đúng danh mục tài liệu tham khảo (references), giữ nguyên số thứ tự [1], [2] ở đầu mỗi tài liệu.

    NỘI DUNG BÀI BÁO THÔ:
    {raw_text}
    """
    
    # Sử dụng gemini-2.5-flash để trích xuất nhanh chóng và chính xác
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(
        prompt,
        generation_config={"response_mime_type": "application/json"}
    )
    
    # Parse kết quả JSON
    try:
        data = json.loads(response.text)
        return data
    except Exception as e:
        # Nếu có lỗi parse, cố gắng làm sạch chuỗi
        clean_text = response.text.strip()
        if clean_text.startswith("```json"):
            clean_text = clean_text[7:]
        if clean_text.endswith("```"):
            clean_text = clean_text[:-3]
        return json.loads(clean_text.strip())

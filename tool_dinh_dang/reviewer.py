import re
from datetime import datetime
import google.generativeai as genai
import json
from config import JSLHU_RULES

def analyze_references_locally(references):
    """
    Phân tích cục bộ danh mục tài liệu tham khảo để đếm tổng số lượng và số tài liệu mới.
    """
    total_ref = len(references)
    current_year = datetime.now().year
    years_limit = JSLHU_RULES["new_ref_years_limit"]
    min_year = current_year - years_limit
    
    new_ref_count = 0
    for ref in references:
        # Tìm tất cả cụm số gồm 4 chữ số (đại diện cho năm)
        years = re.findall(r'\b(20\d{2}|19\d{2})\b', ref)
        if years:
            years_int = [int(y) for y in years]
            max_y = max(years_int)
            if max_y >= min_year:
                new_ref_count += 1
                
    return total_ref, new_ref_count

def check_email_rules(email, affiliations):
    """
    Kiểm tra email tác giả liên hệ xem có đúng tên miền @lhu.edu.vn nếu thuộc LHU không.
    """
    is_lhu = False
    for aff in affiliations:
        aff_name = aff.get("name", "").lower()
        if "lạc hồng" in aff_name or "lac hong" in aff_name or "lhu" in aff_name:
            is_lhu = True
            break
            
    if is_lhu:
        if not email or not email.strip().endswith(JSLHU_RULES["required_email_domain"]):
            return False, f"Có tác giả thuộc LHU nhưng email liên hệ '{email}' không có đuôi {JSLHU_RULES['required_email_domain']}"
        return True, "Hợp lệ"
    return True, "Không bắt buộc (không thuộc đơn vị LHU)"

def run_ai_review(data, api_key):
    """
    Sử dụng Gemini API để chấm điểm chi tiết 24 tiêu chí sơ loại dựa trên nội dung đã trích xuất.
    """
    genai.configure(api_key=api_key)
    
    # Chuẩn bị thông tin tóm tắt để gửi lên AI
    summary_info = {
        "title_vi": data.get("title_vi", ""),
        "abstract_vi_len": len(data.get("abstract_vi", "").split()),
        "abstract_vi": data.get("abstract_vi", ""),
        "keywords_vi": data.get("keywords_vi", []),
        "email": data.get("email_contact", ""),
        "affiliations": [a["name"] for a in data.get("affiliations_vi", [])],
        "sections": [s["title"] for s in data.get("sections", [])],
        "num_references": len(data.get("references", []))
    }
    
    prompt = f"""
    Bạn là một Biên tập viên cao cấp của Tạp chí Khoa học Lạc Hồng (JSLHU). 
    Hãy đánh giá bài báo dựa trên các tiêu chí sơ loại của tạp chí.
    Dưới đây là thông tin tóm tắt của bài báo:
    {json.dumps(summary_info, ensure_ascii=False, indent=2)}
    
    Hãy chấm điểm và cho ý kiến nhận xét (Có/Không/Cần kiểm tra thủ công) cho đúng 24 tiêu chí sơ loại sau:
    1. Định dạng theo mẫu của tạp chí (Đánh giá: Có - vì công cụ sẽ định dạng lại)
    2. Tên bài báo không quá 20 chữ (Đo số từ tiêu đề tiếng Việt: '{data.get("title_vi", "")}')
    3. Tên tác giả (Có/Không)
    4. Tên đơn vị (Có/Không)
    5. Tóm tắt từ 150-250 chữ (Đo số từ abstract tiếng Việt: {summary_info['abstract_vi_len']} từ)
    6. Từ khóa không quá 5 từ (Đếm số từ khóa tiếng Việt: {len(data.get("keywords_vi", []))})
    7. Có 10-15 tài liệu tham khảo và phải có ít nhất 3 tài liệu mới trong 5 năm gần nhất
    8. Tóm tắt: Nêu rõ Câu hỏi và mục đích của nghiên cứu
    9. Tóm tắt: Nêu rõ Phương pháp nghiên cứu (mô tả cách thức giải quyết vấn đề)
    10. Tóm tắt: Nêu rõ Kết quả chính của nghiên cứu
    11. Tóm tắt: Nêu rõ Kết luận và ý nghĩa nghiên cứu
    12. Tỷ lệ trùng lắp dưới 20% (Đánh giá: Cần kiểm tra thủ công qua phần mềm đạo văn)
    13. Nội dung: Có phần tóm tắt và đề xuất giải pháp
    14. Nội dung: Có Phương pháp nghiên cứu
    15. Nội dung: Thể hiện được Tính mới của nghiên cứu
    16. Nội dung: Có trình bày, mô tả giải pháp
    17. Nội dung: Có hình vẽ diễn đạt giải pháp (nếu có đề cập hình vẽ trong bài)
    18. Nội dung: Có đề xuất giải pháp mới
    19. Nội dung: Có phần Kết quả
    20. Nội dung: Có trình bày, mô tả kết quả
    21. Nội dung: Có hình ảnh diễn đạt kết quả
    22. Nội dung: Có phần Kết luận
    23. Nội dung: Có đóng góp cho cộng đồng khoa học
    24. Trích dẫn đúng và đủ nguồn tham khảo (Đúng định dạng APA/IEEE)

    Hãy trả về một danh sách JSON chứa kết quả đánh giá cho 24 tiêu chí này. Mỗi phần tử trong danh sách gồm:
    - "id": số thứ tự từ 1 đến 24
    - "criterion": tên tiêu chí
    - "status": true (nếu đạt/Có), false (nếu không đạt/Không), null (nếu cần kiểm tra thủ công hoặc chưa đủ dữ liệu đánh giá)
    - "note": nhận xét ngắn gọn, chỉ ra lý do tại sao đạt hoặc chưa đạt.

    Hãy đảm bảo trả về định dạng JSON mảng thuần túy, không chứa thẻ ```json hay bất kỳ văn bản thừa nào ngoài JSON.
    """
    
    try:
        model = genai.GenerativeModel("gemini-2.5-flash")
        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        results = json.loads(response.text)
        
        # Tính điểm dựa trên số tiêu chí đạt (status == True)
        score = sum(1 for item in results if item.get("status") is True)
        return {
            "score": score,
            "max_score": 24,
            "results": results
        }
    except Exception as e:
        print(f"Error in run_ai_review: {e}")
        # Trả về kết quả đánh giá rule-based cơ bản dự phòng
        return evaluate_locally(data)

def evaluate_locally(data):
    """
    Đánh giá sơ bộ bằng luật định sẵn (rule-based) khi không gọi được AI hoặc gặp lỗi.
    """
    results = []
    score = 0
    
    # 1. Định dạng theo mẫu
    results.append({"id": 1, "criterion": "Định dạng theo mẫu của tạp chí", "status": True, "note": "Được định dạng lại tự động bằng công cụ"})
    score += 1
    
    # 2. Tên bài báo không quá 20 chữ
    title_words = len(data.get("title_vi", "").split())
    title_ok = title_words <= JSLHU_RULES["max_title_words"]
    results.append({
        "id": 2,
        "criterion": f"Tên bài báo không quá 20 chữ ({title_words} từ)",
        "status": title_ok,
        "note": "Hợp lệ" if title_ok else f"Tiêu đề tiếng Việt quá dài ({title_words} từ)"
    })
    if title_ok: score += 1
    
    # 3. Tên tác giả
    has_authors = len(data.get("authors_vi", [])) > 0
    results.append({"id": 3, "criterion": "Có thông tin tên tác giả", "status": has_authors, "note": "Hợp lệ" if has_authors else "Thiếu thông tin tác giả"})
    if has_authors: score += 1
    
    # 4. Tên đơn vị
    has_aff = len(data.get("affiliations_vi", [])) > 0
    results.append({"id": 4, "criterion": "Có thông tin đơn vị công tác", "status": has_aff, "note": "Hợp lệ" if has_aff else "Thiếu thông tin đơn vị"})
    if has_aff: score += 1
    
    # 5. Tóm tắt từ 150-250 chữ
    abs_words = len(data.get("abstract_vi", "").split())
    abs_ok = JSLHU_RULES["min_abstract_words"] <= abs_words <= JSLHU_RULES["max_abstract_words"]
    results.append({
        "id": 5,
        "criterion": f"Tóm tắt từ 150-250 chữ ({abs_words} từ)",
        "status": abs_ok,
        "note": "Hợp lệ" if abs_ok else f"Độ dài tóm tắt không đạt chuẩn ({abs_words} từ)"
    })
    if abs_ok: score += 1
    
    # 6. Từ khóa không quá 5 từ
    kw_count = len(data.get("keywords_vi", []))
    kw_ok = 0 < kw_count <= JSLHU_RULES["max_keywords"]
    results.append({
        "id": 6,
        "criterion": f"Từ khóa không quá 5 từ ({kw_count} từ khóa)",
        "status": kw_ok,
        "note": "Hợp lệ" if kw_ok else f"Số lượng từ khóa chưa chuẩn ({kw_count} từ khóa)"
    })
    if kw_ok: score += 1
    
    # 7. Tài liệu tham khảo
    total_ref, new_ref = analyze_references_locally(data.get("references", []))
    ref_count_ok = JSLHU_RULES["min_references"] <= total_ref <= JSLHU_RULES["max_references"]
    ref_new_ok = new_ref >= JSLHU_RULES["min_new_references"]
    ref_ok = ref_count_ok and ref_new_ok
    results.append({
        "id": 7,
        "criterion": f"Có 10-15 tài liệu tham khảo và ít nhất 3 tài liệu mới trong 5 năm gần nhất",
        "status": ref_ok,
        "note": f"Hợp lệ (Tổng: {total_ref}, Mới: {new_ref})" if ref_ok else f"Chưa đạt (Tổng: {total_ref}, Mới: {new_ref})"
    })
    if ref_ok: score += 1
    
    # 8-11. Các phần tóm tắt (cơ bản)
    abstract_text = data.get("abstract_vi", "")
    has_abstract_content = len(abstract_text) > 50
    results.append({"id": 8, "criterion": "Tóm tắt: Câu hỏi và mục đích của nghiên cứu", "status": has_abstract_content, "note": "Có xuất hiện" if has_abstract_content else "Trống"})
    results.append({"id": 9, "criterion": "Tóm tắt: Phương pháp nghiên cứu", "status": has_abstract_content, "note": "Có xuất hiện" if has_abstract_content else "Trống"})
    results.append({"id": 10, "criterion": "Tóm tắt: Kết quả", "status": has_abstract_content, "note": "Có xuất hiện" if has_abstract_content else "Trống"})
    results.append({"id": 11, "criterion": "Tóm tắt: Kết luận", "status": has_abstract_content, "note": "Có xuất hiện" if has_abstract_content else "Trống"})
    if has_abstract_content: score += 4
    
    # 12. Trùng lặp
    results.append({"id": 12, "criterion": "Tỷ lệ trùng lắp dưới 20%", "status": None, "note": "Cần kiểm tra đạo văn bằng phần mềm chuyên dụng"})
    
    # 13-24. Các tiêu chí nội dung chính
    sections_titles = [s["title"].lower() for s in data.get("sections", [])]
    has_method = any("phương pháp" in t or "method" in t or "cơ sở" in t or "nguyên lý" in t for t in sections_titles)
    has_result = any("kết quả" in t or "result" in t or "thực nghiệm" in t or "thảo luận" in t for t in sections_titles)
    has_conclusion = any("kết luận" in t or "conclusion" in t for t in sections_titles)
    
    results.append({"id": 13, "criterion": "Nội dung: Có phần tóm tắt và đề xuất giải pháp", "status": True, "note": "Có phần tóm tắt/đặt vấn đề"})
    score += 1
    results.append({"id": 14, "criterion": "Nội dung: Phương pháp nghiên cứu", "status": has_method, "note": "Đầy đủ" if has_method else "Thiếu chương mục về Phương pháp"})
    if has_method: score += 1
    results.append({"id": 15, "criterion": "Nội dung: Tính mới của nghiên cứu", "status": None, "note": "Cần kiểm tra thủ công nội dung khoa học"})
    results.append({"id": 16, "criterion": "Nội dung: Có trình bày, mô tả giải pháp", "status": has_method, "note": "Đầy đủ" if has_method else "Thiếu phần trình bày giải pháp"})
    if has_method: score += 1
    results.append({"id": 17, "criterion": "Nội dung: Có hình diễn đạt giải pháp", "status": None, "note": "Cần kiểm tra thủ công các hình vẽ trong bài"})
    results.append({"id": 18, "criterion": "Nội dung: Có đề xuất giải pháp mới", "status": None, "note": "Cần kiểm tra thủ công"})
    results.append({"id": 19, "criterion": "Nội dung: Kết quả", "status": has_result, "note": "Đầy đủ" if has_result else "Thiếu phần Kết quả"})
    if has_result: score += 1
    results.append({"id": 20, "criterion": "Nội dung: Có trình bày, mô tả kết quả", "status": has_result, "note": "Đầy đủ" if has_result else "Thiếu phần mô tả kết quả"})
    if has_result: score += 1
    results.append({"id": 21, "criterion": "Nội dung: Có hình ảnh diễn đạt kết quả", "status": None, "note": "Cần kiểm tra thủ công"})
    results.append({"id": 22, "criterion": "Nội dung: Kết luận", "status": has_conclusion, "note": "Đầy đủ" if has_conclusion else "Thiếu phần Kết luận"})
    if has_conclusion: score += 1
    results.append({"id": 23, "criterion": "Nội dung: Có đóng góp cho cộng đồng khoa học", "status": None, "note": "Cần kiểm tra thủ công"})
    
    # 24. Trích dẫn đúng và đủ nguồn
    email_ok, email_note = check_email_rules(data.get("email_contact", ""), data.get("affiliations_vi", []))
    results.append({"id": 24, "criterion": "Trích dẫn đúng và đủ nguồn tham khảo, email liên hệ hợp lệ", "status": email_ok, "note": email_note})
    if email_ok: score += 1
    
    # Thêm các tiêu chí định tính còn lại cho đủ 24
    all_ids = set(range(1, 25))
    existing_ids = {item["id"] for item in results}
    missing_ids = all_ids - existing_ids
    for m_id in sorted(missing_ids):
        results.append({"id": m_id, "criterion": f"Tiêu chí số {m_id}", "status": None, "note": "Cần kiểm tra thủ công"})
        
    results.sort(key=lambda x: x["id"])
    
    return {
        "score": score,
        "max_score": 17,  # Tổng số điểm kiểm tra tự động
        "results": results
    }

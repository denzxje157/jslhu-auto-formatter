import docx
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import re
import copy
import io
import os
import sys

def run_tool_bm01(input_file, output_file="output_chuan_bm01_jslhu.docx"):
    """
    TOOL 2: Định dạng Chuẩn Bài Báo Mẫu BM01 - Tạp chí Lạc Hồng JSLHU 
    (Trang bìa Metadata 2 cột có khung viền 0.5pt, Thân bài 2 Cột 1.0cm Gutter, A4, Lề 2.5cm, Font 10pt)
    """
    sys.stdout.reconfigure(encoding='utf-8')
    print(f"\n================ [TOOL 2] BẮT ĐẦU ĐỊNH DẠNG CHUẨN BM01 - JSLHU ================")
    print(f"File đầu vào: '{input_file}'")
    
    if not os.path.exists(input_file):
        print(f"Lỗi: Không tìm thấy file '{input_file}'!")
        return False

    src_doc = docx.Document(input_file)
    out_doc = docx.Document()
    
    # 1. Cấu hình Page Setup A4 cho Section 0 (Metadata trang bìa 1 cột)
    sec0 = out_doc.sections[0]
    sec0.top_margin = Cm(2.5)
    sec0.bottom_margin = Cm(2.5)
    sec0.left_margin = Cm(2.5)
    sec0.right_margin = Cm(2.5)
    sec0.header_distance = Cm(1.2)
    sec0.footer_distance = Cm(1.2)
    
    sectPr0 = sec0._sectPr
    cols0 = sectPr0.xpath('./w:cols')
    if cols0:
        sectPr0.remove(cols0[0])
    sectPr0.append(parse_xml(f'<w:cols {nsdecls("w")} w:num="1" w:space="0"/>'))

    # 2. Copy Bảng Metadata (Bảng 0 & Bảng 1) sang Trang 1
    for t_idx in range(min(2, len(src_doc.tables))):
        t_xml = copy.deepcopy(src_doc.tables[t_idx]._tbl)
        out_doc._element.body.append(t_xml)
        
    # Áp dụng định dạng chuẩn Mẫu BM01 cho Bảng Metadata (Đường viền 0.5pt sz="4")
    for t_idx in range(min(2, len(out_doc.tables))):
        table = out_doc.tables[t_idx]
        table.style = 'Normal Table'
        table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        table.allow_autofit = True
        
        tblPr = table._tbl.tblPr
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)
            
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
            f'  <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideH w:val="none"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)
        
        if len(table.rows) > 0:
            for cell in table.rows[0].cells[:2]:
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>'))
                
        if len(table.rows) > 5:
            for cell in table.rows[5].cells[:2]:
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>'))

        if len(table.rows) > 6:
            for c_idx, cell in enumerate(table.rows[6].cells):
                tcPr = cell._tc.get_or_add_tcPr()
                if c_idx < 2:
                    tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>'))
                else:
                    tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="nil"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>'))

    # 3. Tạo Section 2 Cột cho Thân bài BM01 (Gutter = 1.0cm / 567 dxa)
    sec_body = out_doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    sec_body.top_margin = Cm(2.5)
    sec_body.bottom_margin = Cm(2.5)
    sec_body.left_margin = Cm(2.5)
    sec_body.right_margin = Cm(2.5)
    
    sectPr = sec_body._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        sectPr.remove(cols[0])
    cols_xml = parse_xml(f'<w:cols {nsdecls("w")} w:num="2" w:space="567"/>')
    sectPr.append(cols_xml)

    # 4. Trích xuất văn bản thân bài từ src_doc sang Section 2 (Bên trong 2 Cột BM01)
    start_idx = 0
    for idx, p in enumerate(src_doc.paragraphs):
        text = p.text.strip()
        if text.startswith("1. ") or text.startswith("1. Giới thiệu") or text.startswith("1. Gioi thieu"):
            start_idx = idx
            break

    images_bytes = []
    for item in src_doc.element.body:
        drawings = item.xpath('.//w:drawing')
        if drawings:
            for dr in drawings:
                embed_ids = dr.xpath('.//a:blip/@r:embed')
                if embed_ids:
                    rId = embed_ids[0]
                    try:
                        image_part = src_doc.part.related_parts[rId]
                        images_bytes.append(image_part.blob)
                    except Exception:
                        pass

    img_idx = 0
    ref_started = False

    for idx in range(start_idx, len(src_doc.paragraphs)):
        p_src = src_doc.paragraphs[idx]
        text = p_src.text
        
        if not text.strip() and not p_src._element.xpath('.//*[local-name()="oMath"]') and not p_src._element.xpath('.//w:drawing'):
            continue
            
        if (text.strip().startswith("Hình ") or text.strip().startswith("Figure ") or text.strip().startswith("Hinh ")) and ":" in text:
            if img_idx < len(images_bytes):
                img_p = out_doc.add_paragraph()
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_p.paragraph_format.space_before = Pt(4)
                img_p.paragraph_format.space_after = Pt(2)
                img_p.paragraph_format.first_line_indent = Cm(0)
                
                run_img = img_p.add_run()
                run_img.add_picture(io.BytesIO(images_bytes[img_idx]), width=Inches(3.1))
                img_idx += 1
                
            p_cap = out_doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(4)
            p_cap.paragraph_format.first_line_indent = Cm(0)
            r_cap = p_cap.add_run(text.strip())
            r_cap.font.name = 'Times New Roman'
            r_cap.font.size = Pt(9.5)
            r_cap.italic = True
            r_cap.bold = True
            continue

        p_new = out_doc.add_paragraph()
        p_fmt = p_new.paragraph_format
        
        is_heading = False
        if re.match(r'^\d+(\.\d+)*\.', text.strip()):
            is_heading = True
        elif text.strip().startswith("5. ") or "tài liệu tham khảo" in text.lower():
            is_heading = True
            ref_started = True
            
        if is_heading:
            p_fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_fmt.first_line_indent = Cm(0)
            p_fmt.space_before = Pt(6)
            p_fmt.space_after = Pt(4)
            p_fmt.line_spacing = 1.0
            
            run = p_new.add_run(text.strip().upper() if not re.search(r'\d+\.\d+', text.strip()) else text.strip())
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            run.bold = True
        elif ref_started:
            p_fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_fmt.left_indent = Cm(0.5)
            p_fmt.first_line_indent = -Cm(0.5)
            p_fmt.space_before = Pt(0)
            p_fmt.space_after = Pt(2)
            p_fmt.line_spacing = 1.0
            
            for r_src in p_src.runs:
                if not r_src.text:
                    continue
                run = p_new.add_run(r_src.text)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(9)
                run.bold = r_src.bold
                run.italic = r_src.italic
        else:
            p_fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_fmt.first_line_indent = Cm(0.5)
            p_fmt.space_before = Pt(0)
            p_fmt.space_after = Pt(3)
            p_fmt.line_spacing = 1.0
            
            for r_src in p_src.runs:
                if not r_src.text:
                    continue
                run = p_new.add_run(r_src.text)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(10)
                run.bold = r_src.bold
                run.italic = r_src.italic

    # 5. Copy & Nhận diện BẢNG BIỂU NỘI DUNG -> Căn giữa 2 Cột chuẩn BM01
    for t_src in src_doc.tables[2:]:
        t_xml = copy.deepcopy(t_src._tbl)
        out_doc._element.body.append(t_xml)
        
    for t in out_doc.tables[2:]:
        t.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        t.allow_autofit = True
        
        tblPr = t._tbl.tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is not None:
            tblPr.remove(tblW)
        tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="0" w:type="auto"/>'))
        
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
            f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideH w:val="none"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)
        
        if len(t.rows) > 0:
            for cell in t.rows[0].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcBorders = tcPr.find(qn('w:tcBorders'))
                if tcBorders is not None:
                    tcPr.remove(tcBorders)
                tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>'))

        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.0
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(9)

    out_doc.save(output_file)
    print(f"-> TOOL 2 hoàn thành xuất file CHUẨN BM01 2 CỘT: '{output_file}'!")
    return True

if __name__ == "__main__":
    inp = sys.argv[1] if len(sys.argv) > 1 else "bao_cao_nghien_cuu_RAG_hoan_chinh_v4.docx"
    run_tool_bm01(inp)

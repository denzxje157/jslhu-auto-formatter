import docx
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import os
import sys

def run_tool_bm01(input_file, output_file="output_chuan_bm01_jslhu.docx"):
    """
    TOOL 2: Định dạng Chuẩn Bài Báo Mẫu BM01 - Tạp chí Lạc Hồng JSLHU 
    (Trang bìa Metadata 2 cột có khung viền 0.5pt, Thân bài 2 Cột 1cm gutter, A4, Lề 2.5cm, 10pt)
    """
    sys.stdout.reconfigure(encoding='utf-8')
    print(f"\n================ [TOOL 2] BẮT ĐẦU ĐỊNH DẠNG CHUẨN BM01 - JSLHU ================")
    print(f"File đầu vào: '{input_file}'")
    
    if not os.path.exists(input_file):
        print(f"Lỗi: Không tìm thấy file '{input_file}'!")
        return False

    doc = docx.Document(input_file)
    
    # 1. Cấu hình Page Setup cho Section 0 (Section đầu chứa metadata & thân bài 2 cột)
    sec0 = doc.sections[0]
    sec0.top_margin = Cm(2.5)
    sec0.bottom_margin = Cm(2.5)
    sec0.left_margin = Cm(2.5)
    sec0.right_margin = Cm(2.5)
    sec0.header_distance = Cm(1.2)
    sec0.footer_distance = Cm(1.2)
    
    # 2. Cấu hình Bảng Metadata (Bảng 0 & Bảng 1) chuẩn đường kẻ 0.5pt (sz="4")
    for t_idx in range(min(2, len(doc.tables))):
        table = doc.tables[t_idx]
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
        
        # Đường kẻ viền 0.5pt (sz="4")
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

    doc.save(output_file)
    print(f"-> Đã xuất file chuẩn BM01 thành công: '{output_file}'!")
    return True

if __name__ == "__main__":
    inp = sys.argv[1] if len(sys.argv) > 1 else "bao_cao_nghien_cuu_RAG_hoan_chinh_v4.docx"
    run_tool_bm01(inp)

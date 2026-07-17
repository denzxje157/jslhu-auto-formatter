import sys
import os

def run_both():
    sys.stdout.reconfigure(encoding='utf-8')
    print("==========================================================================")
    print("       CHƯƠNG TRÌNH ĐỊNH DẠNG TỰ ĐỘNG BÀI BÁO TẠP CHÍ LẠC HỒNG (JSLHU)")
    print("==========================================================================\n")
    
    inp_file = sys.argv[1] if len(sys.argv) > 1 else "bao_cao_nghien_cuu_RAG_hoan_chinh_v4.docx"
    
    if not os.path.exists(inp_file):
        print(f"Lỗi: Không tìm thấy tệp đầu vào '{inp_file}'!")
        return

    # 1. Chạy Tool 1: Vòng sơ loại (1 Cột, A4, 11pt, lề 3-2-3-2cm)
    from tool_dinh_dang.tool_so_loai import run_tool_so_loai
    run_tool_so_loai(inp_file, "bao_cao_nghien_cuu_RAG_ban_nop_so_loai.docx")
    
    # 2. Chạy Tool 2: Chuẩn BM01 (2 Cột, A4, 10pt, lề 2.5cm, Khung viền 0.5pt)
    from tool_dinh_dang.tool_bm01 import run_tool_bm01
    run_tool_bm01(inp_file, "bao_cao_nghien_cuu_RAG_hoan_chinh_v4.docx")
    
    print("\n--------------------------------------------------------------------------")
    print(" HOÀN THÀNH TẤT CẢ! CẢ 2 TOOL ĐÃ TẠO XONG FILE XUẤT NGUYÊN BẢN CHUẨN ĐẸP 100%.")
    print("--------------------------------------------------------------------------")

if __name__ == "__main__":
    run_both()

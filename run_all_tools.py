import sys
import os

def run_both():
    sys.stdout.reconfigure(encoding='utf-8')
    print("==========================================================================")
    print("       CHƯƠNG TRÌNH ĐỊNH DẠNG TỰ ĐỘNG BÀI BÁO TẠP CHÍ LẠC HỒNG (JSLHU)")
    print("==========================================================================\n")
    
    inp_file = sys.argv[1] if len(sys.argv) > 1 else "bao_cao_nghien_cuu_RAG_hoan_chinh_tieng_viet.docx"
    
    if not os.path.exists(inp_file):
        print(f"Lỗi: Không tìm thấy tệp đầu vào '{inp_file}'!")
        return

    # 1. Chạy Tool 1: Vòng sơ loại -> Xuất ra file mới riêng biệt: output_ban_nop_so_loai.docx
    from tool_dinh_dang.tool_so_loai import run_tool_so_loai
    run_tool_so_loai(inp_file, "output_ban_nop_so_loai.docx")
    
    # 2. Chạy Tool 2: Chuẩn BM01 -> Xuất ra file mới riêng biệt: output_chuan_bm01_jslhu.docx
    from tool_dinh_dang.tool_bm01 import run_tool_bm01
    run_tool_bm01(inp_file, "output_chuan_bm01_jslhu.docx")
    
    print("\n--------------------------------------------------------------------------")
    print(" HOÀN THÀNH TẤT CẢ! FILE KẾT QUẢ MỚI ĐÃ ĐƯỢC SINH RA RIÊNG BIỆT:")
    print("  - [Tool 1 Sơ loại]: output_ban_nop_so_loai.docx")
    print("  - [Tool 2 BM01]:    output_chuan_bm01_jslhu.docx")
    print("--------------------------------------------------------------------------")

if __name__ == "__main__":
    run_both()

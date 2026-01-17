import fitz  # PyMuPDF
import os
import sys

def xu_ly_pdf_A4_A3_Triet_De():
    # 1. XÁC ĐỊNH ĐƯỜNG DẪN CHUẨN (Quan trọng để không bị file trống)
    if getattr(sys, 'frozen', False):
        curr_dir = os.path.dirname(sys.executable)
    else:
        curr_dir = os.path.dirname(os.path.abspath(__file__))
    
    output_folder = os.path.join(curr_dir, "KET_QUA_GOM_FILE")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 2. QUÉT FILE (Hỗ trợ tên tiếng Nhật và ký tự đặc biệt)
    files = [f for f in os.listdir(curr_dir) if f.lower().endswith(".pdf")]
    
    if not files:
        print(f"KHÔNG THẤY FILE PDF NÀO TẠI: {curr_dir}")
        return

    # Sắp xếp A4 trước, A3 sau
    files.sort(key=lambda x: ("(A4)" not in x, "(A3)" not in x))

    page_counter = 1
    blue = (0, 0, 1)

    for filename in files:
        print(f"Đang xử lý: {filename}")
        doc = fitz.open(os.path.join(curr_dir, filename))
        new_doc = fitz.open() # Tạo file mới để thực hiện Shrink (thu nhỏ)

        for page in doc:
            # XOAY: Nếu file A4 trang ngang -> Xoay ngược chiều kim đồng hồ (CCW)
            if "(A4)" in filename and page.rect.width > page.rect.height:
                page.set_rotation((page.rotation + 90) % 360)
            
            rect = page.rect
            w, h = rect.width, rect.height

            # CƠ CHẾ SHRINK: Tạo lề trắng 5% ở dưới để đánh số không bị đè
            new_page = new_doc.new_page(width=w, height=h)
            # Thu nhỏ nội dung 95% và đẩy lên trên
            shrink_rect = fitz.Rect(w*0.02, h*0.01, w*0.98, h*0.95) 
            new_page.show_pdf_page(shrink_rect, doc, page.number)

            # CHỌN VỊ TRÍ A, B, C (Tọa độ sau khi đã Shrink)
            if w < 610:  # A4 Dọc (Vị trí A)
                pos = fitz.Point(w - 55, h - 30)
                f_size = 11
            elif w < 900 and w < h:  # A3 Dọc (Vị trí B)
                pos = fitz.Point(w - 75, h - 40)
                f_size = 13
            elif w > 1100 and w > h:  # A3 Ngang (Vị trí C)
                pos = fitz.Point(w - 110, h - 55)
                f_size = 14
            else:  # Khác
                pos = fitz.Point(w * 0.94, h * 0.97)
                f_size = h * 0.015

            # GHI SỐ TRANG
            new_page.insert_text(
                pos, f"Page {page_counter}",
                fontsize=f_size, fontname="helv", color=blue,
                align=fitz.TEXT_ALIGN_RIGHT, overlay=True
            )
            page_counter += 1

        output_path = os.path.join(output_folder, f"Checked_{filename}")
        new_doc.save(output_path)
        new_doc.close()
        doc.close()

    print(f"\n--- XONG! Kiểm tra thư mục KET_QUA_GOM_FILE ---")

if __name__ == "__main__":
    xu_ly_pdf_A4_A3_Triet_De()

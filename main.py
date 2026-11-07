import streamlit as st
import os
import re
import random
import zipfile
import io
import copy
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import pandas as pd

# 1. CẤU HÌNH TRANG WEB STREAMLIT
st.set_page_config(page_title="Phần mềm Trộn Đề Vật Lí Chuyên Nghiệp", page_icon="📝", layout="centered")

st.title("📝 Ứng Dụng Trộn Đề Vật Lí Online")
st.caption("👨‍🏫 **Tác giả:** LQChau - CLB STEM - THPT Thủ Đức")
st.write("Hệ thống trộn đề thi chuẩn cấu trúc Bộ GD&ĐT, bảo toàn định dạng font/size gốc, công thức MathType, Equation, Hình ảnh & Bảng.")

st.divider()

# 2. HÀM CHÈN KHỐI XML WORD
def append_elem(doc, elem):
    body = doc._element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(elem)
    else:
        body.append(elem)

def Kiem_Tra_Gach_Chan(paragraph):
    for r in paragraph.runs:
        if r.text and r.text.strip():
            if bool(r.underline) and r.underline != 0 and r.underline is not False:
                return True
    return False

# 3. HÀM CẬP NHẬT TIỀN TỐ CÂU & PHƯƠNG ÁN (KẾ THỪA FONT CHỮ/SIZE GỐC)
def Cap_Nhat_Tien_To_Cau(p_elem, new_prefix):
    doc_temp = Document()
    append_elem(doc_temp, copy.deepcopy(p_elem))
    p = doc_temp.paragraphs[0]
    
    # Lấy thông tin font chữ & font size từ run đầu tiên có văn bản
    font_name, font_size = None, None
    for r in p.runs:
        if r.text and r.text.strip():
            font_name = r.font.name
            font_size = r.font.size
            break
            
    full_text = p.text
    match = re.match(r'^(Câu|Cau)\s*\d+[\.\:\-\)\s]*', full_text, re.IGNORECASE)
    if match:
        prefix_len = len(match.group(0))
        remains = prefix_len
        for r in p.runs:
            if not r.text:
                continue
            if len(r.text) <= remains:
                remains -= len(r.text)
                r.text = ""
            else:
                r.text = r.text[remains:]
                remains = 0
                break
                
    r0 = p.add_run(new_prefix)
    r0.bold = True
    if font_name: r0.font.name = font_name
    if font_size: r0.font.size = font_size
    p._element.insert(0, r0._element)
    return p._element

def Lay_Noi_Dung_Phuong_An_Clean(p_elem, new_prefix):
    doc_temp = Document()
    append_elem(doc_temp, copy.deepcopy(p_elem))
    p = doc_temp.paragraphs[0]
    
    font_name, font_size = None, None
    for r in p.runs:
        if r.text and r.text.strip():
            font_name = r.font.name
            font_size = r.font.size
            break
            
    full_text = p.text
    match = re.match(r'^[A-Da-d][\.\:\-\)\s]*', full_text)
    if match:
        prefix_len = len(match.group(0))
        remains = prefix_len
        for r in p.runs:
            if not r.text:
                continue
            if len(r.text) <= remains:
                remains -= len(r.text)
                r.text = ""
            else:
                r.text = r.text[remains:]
                remains = 0
                break
                
    r0 = p.add_run(new_prefix)
    r0.bold = True
    if font_name: r0.font.name = font_name
    if font_size: r0.font.size = font_size
    p._element.insert(0, r0._element)
    return p._element

# 4. HÀM DÀN HÀNG ĐÁP ÁN PHẦN I BẰNG TAB (TĂNG TÍNH THẨM MỸ)
def Render_Phuong_An_Phan_1(doc, opts, nhan_p1):
    total_len = sum(len(opt["text"]) for opt in opts)
    max_len = max(len(opt["text"]) for opt in opts) if opts else 0
    
    # Kịch bản 1: 4 phương án trên 1 dòng (nếu ngắn)
    if total_len < 45 and max_len < 15:
        p_combined = Document().add_paragraph()
        p_format = p_combined.paragraph_format
        p_format.tab_stops.add_tab_stop(Cm(4.3))
        p_format.tab_stops.add_tab_stop(Cm(8.7))
        p_format.tab_stops.add_tab_stop(Cm(13.1))
        
        for idx, opt in enumerate(opts):
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(opt["elem"]), prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            for child in doc_tmp.paragraphs[0]._element:
                p_combined._element.append(copy.deepcopy(child))
            if idx < 3:
                p_combined.add_run("\t")
                
        append_elem(doc, p_combined._element)
        
    # Kịch bản 2: 2 phương án / 1 dòng
    elif total_len < 90 and max_len < 30:
        p_line1 = Document().add_paragraph()
        p_line1.paragraph_format.tab_stops.add_tab_stop(Cm(8.7))
        for idx in [0, 1]:
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(opts[idx]["elem"]), prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            for child in doc_tmp.paragraphs[0]._element:
                p_line1._element.append(copy.deepcopy(child))
            if idx == 0:
                p_line1.add_run("\t")
        append_elem(doc, p_line1._element)
        
        p_line2 = Document().add_paragraph()
        p_line2.paragraph_format.tab_stops.add_tab_stop(Cm(8.7))
        for idx in [2, 3]:
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(opts[idx]["elem"]), prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            for child in doc_tmp.paragraphs[0]._element:
                p_line2._element.append(copy.deepcopy(child))
            if idx == 2:
                p_line2.add_run("\t")
        append_elem(doc, p_line2._element)
        
    # Kịch bản 3: Mỗi phương án 1 dòng
    else:
        for idx, opt in enumerate(opts):
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(opt["elem"]), prefix)
            append_elem(doc, elem)

# 5. HÀM ĐỌC VÀ PHÂN TÍCH ĐỀ GỐC
def doc_de_goc(file_stream):
    doc = Document(file_stream)
    phan1, phan2, phan3 = [], [], []
    current_section = None
    current_item = None

    for p in doc.paragraphs:
        line = p.text.strip()
        if not line:
            continue

        if re.search(r'(?:\[?\s*PHAN[_\s]*3\s*\]?|PHẦN\s*(III|3)\b)', line, re.IGNORECASE):
            if current_item:
                if current_section == 1: phan1.append(current_item)
                elif current_section == 2: phan2.append(current_item)
                current_item = None
            current_section = 3
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*2\s*\]?|PHẦN\s*(II|2)\b)', line, re.IGNORECASE):
            if current_item:
                if current_section == 1: phan1.append(current_item)
                elif current_section == 2: phan2.append(current_item)
                current_item = None
            current_section = 2
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*1\s*\]?|PHẦN\s*(I|1)\b)', line, re.IGNORECASE):
            if current_item:
                if current_section == 1: phan1.append(current_item)
                elif current_section == 2: phan2.append(current_item)
                current_item = None
            current_section = 1
            continue

        is_question = bool(re.match(r'^(Câu|Cau)\s*\d+[\.\:\-\)\s]', line, re.IGNORECASE))

        if current_section == 1:
            if is_question:
                if current_item: phan1.append(current_item)
                current_item = {"q_elem": p._element, "options": [], "correct_idx": None}
            elif re.match(r'^[A-Da-d][\.\:\-\)\s]', line) and current_item:
                clean_opt_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', line).strip()
                is_under = Kiem_Tra_Gach_Chan(p)
                opt_idx = len(current_item["options"])
                current_item["options"].append({
                    "elem": p._element,
                    "text": clean_opt_text
                })
                if is_under:
                    current_item["correct_idx"] = opt_idx

        elif current_section == 2:
            if is_question:
                if current_item: phan2.append(current_item)
                current_item = {"q_elem": p._element, "sub_options": []}
            elif re.match(r'^[A-Da-d][\.\:\-\)\s]', line) and current_item:
                clean_sub_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', line).strip()
                is_under = Kiem_Tra_Gach_Chan(p)
                current_item["sub_options"].append({
                    "elem": p._element,
                    "text": clean_sub_text,
                    "is_correct": "Đúng" if is_under else "Sai"
                })

        elif current_section == 3:
            if is_question:
                if current_item: phan3.append(current_item)
                current_item = {"q_elem": p._element, "correct": ""}
            elif re.search(r'Đáp\s*số\s*[\:\-\.]', line, re.IGNORECASE) and current_item:
                current_item["correct"] = re.sub(r'^Đáp\s*số\s*[\:\-\.]\s*', '', line, flags=re.IGNORECASE).strip()

    if current_item:
        if current_section == 1: phan1.append(current_item)
        elif current_section == 2: phan2.append(current_item)
        elif current_section == 3: phan3.append(current_item)

    return phan1, phan2, phan3

# 6. KHU VỰC THIẾT LẬP THÔNG TIN (UI STREAMLIT)
with st.expander("⚙️ **CẤU HÌNH THÔNG TIN TIÊU ĐỀ ĐỀ THI & MÃ ĐỀ**", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        so_gd = st.text_input("Sở Giáo dục & Đào tạo", value="SỞ GIÁO DỤC VÀ ĐÀO TẠO TP.HCM")
        ten_truong = st.text_input("Tên Trường", value="TRƯỜNG THPT THỦ ĐỨC")
        ten_ky_thi = st.text_input("Tên Kỳ Thi", value="KIỂM TRA ĐỊNH KÌ GIỮA HỌC KÌ I")
    with col2:
        mon_thi = st.text_input("Môn thi / Khối", value="MÔN VẬT LÍ - KHỐI 10")
        thoi_gian = st.text_input("Thời gian làm bài", value="45 phút")
        ds_ma_de_input = st.text_input("Danh sách Mã Đề (cách nhau bởi dấu phẩy)", value="2101, 2102, 2103, 2104")

uploaded_file = st.file_uploader("Bước 1: Tải lên file Word đề gốc (.docx)", type=["docx"])

if uploaded_file is not None:
    st.success(f"✅ Đã nhận file: **{uploaded_file.name}**")

if st.button("🚀 Bước 2: Bắt Đầu Trộn Đề & Sinh Đáp Án", type="primary"):
    if uploaded_file is None:
        st.warning("⚠️ Vui lòng chọn file Word đề gốc ở Bước 1 trước!")
    else:
        with st.spinner("Đang khởi tạo mã đề thi chuẩn A4 chuyên nghiệp..."):
            phan1_goc, phan2_goc, phan3_goc = doc_de_goc(uploaded_file)
            
            danh_sach_ma_de = [m.strip() for m in ds_ma_de_input.split(",") if m.strip()]
            if not danh_sach_ma_de:
                danh_sach_ma_de = ["2101", "2102", "2103", "2104"]
                
            tong_so_cau = len(phan1_goc) + len(phan2_goc) + len(phan3_goc)
            if tong_so_cau == 0:
                st.error("❌ **Không tìm thấy câu hỏi!** Vui lòng kiểm tra lại định dạng đề gốc.")
            else:
                st.info(f"📊 Đã phân tích: **{len(phan1_goc)}** câu Phần I | **{len(phan2_goc)}** câu Phần II | **{len(phan3_goc)}** câu Phần III")
                
                zip_buffer = io.BytesIO()
                danh_sach_excel_data = []

                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for ma_de in danh_sach_ma_de:
                        doc = Document()
                        
                        # --- THIẾT LẬP BỐ CỤC TRANG A4 VÀ LỀ TRANG CHUẨN ---
                        sec = doc.sections[0]
                        sec.page_width = Cm(21.0)
                        sec.page_height = Cm(29.7)
                        sec.top_margin = Cm(1.5)
                        sec.bottom_margin = Cm(1.5)
                        sec.left_margin = Cm(2.0)
                        sec.right_margin = Cm(1.5)
                        
                        # --- 1. TẠO BẢNG TIÊU ĐỀ HEADERS (2 CỘT CHUẨN BỘ) ---
                        table = doc.add_table(rows=1, cols=2)
                        table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        table.autofit = False
                        
                        cell_left, cell_right = table.rows[0].cells
                        cell_left.width = Cm(8.5)
                        cell_right.width = Cm(9.0)
                        
                        # Cột trái (Sở & Trường)
                        p_l = cell_left.paragraphs[0]
                        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_so = p_l.add_run(f"{so_gd.upper()}\n")
                        r_so.font.size = Pt(10)
                        r_truong = p_l.add_run(f"{ten_truong.upper()}")
                        r_truong.bold = True
                        r_truong.font.size = Pt(10.5)
                        
                        # Tạo chuỗi ghi chú số câu động thực tế
                        sec_notes = []
                        if len(phan1_goc) > 0: sec_notes.append(f"{len(phan1_goc)} TN")
                        if len(phan2_goc) > 0: sec_notes.append(f"{len(phan2_goc)} ĐS")
                        if len(phan3_goc) > 0: sec_notes.append(f"{len(phan3_goc)} TL Ngắn")
                        note_str = f"({', '.join(sec_notes)})" if sec_notes else ""
                        
                        # Cột phải (Kỳ thi, môn, thời gian, số câu thực tế)
                        p_r = cell_right.paragraphs[0]
                        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_kt = p_r.add_run(f"{ten_ky_thi.upper()}\n")
                        r_kt.bold = True
                        r_kt.font.size = Pt(10.5)
                        
                        r_mon = p_r.add_run(f"Môn thi: {mon_thi.upper()}\n")
                        r_mon.bold = True
                        r_mon.font.size = Pt(10)
                        
                        r_tg = p_r.add_run(f"Thời gian: {thoi_gian}\n")
                        r_tg.font.size = Pt(9.5)
                        r_tg.italic = True
                        
                        r_note = p_r.add_run(f"Đề có {note_str}")
                        r_note.font.size = Pt(9)
                        r_note.italic = True
                        
                        # Dòng thông tin thí sinh
                        p_info = doc.add_paragraph()
                        p_info.paragraph_format.space_before = Pt(8)
                        p_info.paragraph_format.space_after = Pt(12)
                        p_info.paragraph_format.tab_stops.add_tab_stop(Cm(11.5))
                        p_info.paragraph_format.tab_stops.add_tab_stop(Cm(15.0))
                        
                        r_info = p_info.add_run("Họ tên học sinh: .....................................\tSố báo danh: ............................\tMã đề: ")
                        r_info.font.size = Pt(10.5)
                        r_code = p_info.add_run(f"{ma_de}")
                        r_code.bold = True
                        r_code.font.size = Pt(11)

                        nhan_p1 = ["A.", "B.", "C.", "D."]
                        nhan_p2 = ["a)", "b)", "c)", "d)"]

                        # --- 2. TRỘN PHẦN I ---
                        if phan1_goc:
                            p_head1 = doc.add_paragraph()
                            p_head1.paragraph_format.space_before = Pt(6)
                            p_head1.paragraph_format.space_after = Pt(4)
                            r_h1 = p_head1.add_run(f"PHẦN I. Thí sinh trả lời từ câu 1 đến câu {len(phan1_goc)}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
                            r_h1.bold = True
                            r_h1.font.size = Pt(10.5)

                            p1_shuffled = copy.deepcopy(phan1_goc)
                            random.shuffle(p1_shuffled)
                            for idx, q in enumerate(p1_shuffled):
                                updated_q = Cap_Nhat_Tien_To_Cau(q["q_elem"], f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q)
                                
                                opts = q["options"]
                                correct_idx = q["correct_idx"]
                                for o_i, opt in enumerate(opts):
                                    opt["is_correct"] = (o_i == correct_idx)
                                
                                random.shuffle(opts)
                                correct_label = "Chưa chọn đáp án ở đề gốc"
                                for o_idx, opt in enumerate(opts):
                                    if opt["is_correct"]:
                                        correct_label = nhan_p1[o_idx]
                                
                                Render_Phuong_An_Phan_1(doc, opts, nhan_p1)
                                
                                danh_sach_excel_data.append({
                                    "Mã Đề": ma_de, "Phần": "Phần I", "Câu Số": idx + 1, "Đáp Án Đúng": correct_label
                                })

                        # --- 3. TRỘN PHẦN II ---
                        if phan2_goc:
                            p_head2 = doc.add_paragraph()
                            p_head2.paragraph_format.space_before = Pt(10)
                            p_head2.paragraph_format.space_after = Pt(4)
                            r_h2 = p_head2.add_run(f"PHẦN II. Thí sinh trả lời từ câu 1 đến câu {len(phan2_goc)}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
                            r_h2.bold = True
                            r_h2.font.size = Pt(10.5)

                            p2_shuffled = copy.deepcopy(phan2_goc)
                            random.shuffle(p2_shuffled)
                            for idx, q in enumerate(p2_shuffled):
                                updated_q = Cap_Nhat_Tien_To_Cau(q["q_elem"], f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q)
                                
                                subs = q["sub_options"]
                                random.shuffle(subs)
                                for s_idx, sub in enumerate(subs):
                                    # Thẳng lề trái, không dùng tab
                                    prefix = f"{nhan_p2[s_idx]} "
                                    elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(sub["elem"]), prefix)
                                    append_elem(doc, elem)
                                    
                                    danh_sach_excel_data.append({
                                        "Mã Đề": ma_de, "Phần": "Phần II", "Câu Số": f"Câu {idx + 1} ý {nhan_p2[s_idx]}", "Đáp Án Đúng": sub["is_correct"]
                                    })

                        # --- 4. TRỘN PHẦN III ---
                        if phan3_goc:
                            p_head3 = doc.add_paragraph()
                            p_head3.paragraph_format.space_before = Pt(10)
                            p_head3.paragraph_format.space_after = Pt(4)
                            r_h3 = p_head3.add_run(f"PHẦN III. Thí sinh trả lời từ câu 1 đến câu {len(phan3_goc)}.")
                            r_h3.bold = True
                            r_h3.font.size = Pt(10.5)

                            p3_shuffled = copy.deepcopy(phan3_goc)
                            random.shuffle(p3_shuffled)
                            for idx, q in enumerate(p3_shuffled):
                                updated_q = Cap_Nhat_Tien_To_Cau(q["q_elem"], f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q)
                                
                                danh_sach_excel_data.append({
                                    "Mã Đề": ma_de, "Phần": "Phần III", "Câu Số": idx + 1, "Đáp Án Đúng": q["correct"]
                                })

                        # --- 5. LỜI DẶN CHÂN TRANG ---
                        p_foot1 = doc.add_paragraph("---Hết---")
                        p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foot1.paragraph_format.space_before = Pt(12)
                        
                        p_foot2 = doc.add_paragraph("Thí Sinh Không Được Sử Dụng Tài Liệu - Cán Bộ Coi Thi Không Giải Thích Gì Thêm.")
                        p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foot2.runs[0].font.size = Pt(9.5)
                        p_foot2.runs[0].italic = True

                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"{ma_de}.docx", doc_io.getvalue())
                    
                    # --- XUẤT FILE EXCEL ĐÁP ÁN TỔNG HỢP ---
                    df = pd.DataFrame(danh_sach_excel_data)
                    excel_io = io.BytesIO()
                    with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="Dap_An_Tong_Hop")
                    zip_file.writestr("Bang_Dap_An_Cac_Ma_De.xlsx", excel_io.getvalue())

                st.success("🎉 Trộn đề hoàn tất! Các mã đề được định dạng khổ A4 chuyên nghiệp chuẩn 100% Bộ GD&ĐT.")
                st.download_button(
                    label="📥 TẢI XUỐNG TOÀN BỘ ĐỀ THI & FILE EXCEL ĐÁP ÁN (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Bo_De_Thi_Chuyen_Nghiep.zip",
                    mime="application/zip",
                    type="primary"
                )

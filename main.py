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
from docx.oxml import parse_xml
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Alignment, Font
import pandas as pd

# 1. CẤU HÌNH TRANG WEB STREAMLIT
st.set_page_config(page_title="Phần mềm Trộn Đề Vật Lí Chuyên Nghiệp", page_icon="📝", layout="centered")

st.title("📝 Ứng Dụng Trộn Đề Vật Lí Online")
st.caption("👨‍🏫 **Tác giả:** LQChau - CLB STEM - THPT Thủ Đức")
st.write("Hệ thống trộn đề thi chuẩn cấu trúc Bộ GD&ĐT, bảo toàn công thức MathType, Equation, Hình ảnh & xuất file Excel đáp án ma trận chuẩn mực.")

st.divider()

# 2. HÀM KHÓA CỨNG FONT TIMES NEW ROMAN 12PT Ở CẤP ĐỘ MÁY CHỦ
def Thiet_Lap_Font_Times_New_Roman(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    rPr = style._element.get_or_add_rPr()
    rFonts = parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>')
    rPr.append(rFonts)

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

# 3. HÀM CẬP NHẬT TIỀN TỐ CÂU & PHƯƠNG ÁN (GÁN CHẶT TIMES NEW ROMAN 12PT)
def Cap_Nhat_Tien_To_Cau(p_elem, new_prefix):
    doc_temp = Document()
    append_elem(doc_temp, copy.deepcopy(p_elem))
    p = doc_temp.paragraphs[0]
    
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
    r0.font.name = 'Times New Roman'
    r0.font.size = Pt(12)
    p._element.insert(0, r0._element)
    return p._element

def Lay_Noi_Dung_Phuong_An_Clean(p_elem, new_prefix):
    doc_temp = Document()
    append_elem(doc_temp, copy.deepcopy(p_elem))
    p = doc_temp.paragraphs[0]
    
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
    r0.font.name = 'Times New Roman'
    r0.font.size = Pt(12)
    p._element.insert(0, r0._element)
    return p._element

# 4. HÀM DÀN HÀNG ĐÁP ÁN PHẦN I BẰNG TAB TỰ ĐỘNG
def Render_Phuong_An_Phan_1(doc, opts, nhan_p1):
    total_len = sum(len(opt["text"]) for opt in opts)
    max_len = max(len(opt["text"]) for opt in opts) if opts else 0
    
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
                r_t = p_combined.add_run("\t")
                r_t.font.name = 'Times New Roman'
                r_t.font.size = Pt(12)
                
        append_elem(doc, p_combined._element)
        
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
                r_t = p_line1.add_run("\t")
                r_t.font.name = 'Times New Roman'
                r_t.font.size = Pt(12)
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
                r_t = p_line2.add_run("\t")
                r_t.font.name = 'Times New Roman'
                r_t.font.size = Pt(12)
        append_elem(doc, p_line2._element)
        
    else:
        for idx, opt in enumerate(opts):
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(opt["elem"]), prefix)
            append_elem(doc, elem)

# 5. HÀM ĐỌC ĐỀ GỐC
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
        with st.spinner("Đang khởi tạo mã đề thi và tạo file Excel đáp án ma trận..."):
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
                ma_de_answers = {}

                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for ma_de in danh_sach_ma_de:
                        doc = Document()
                        Thiet_Lap_Font_Times_New_Roman(doc)
                        ma_de_answers[ma_de] = {'p1': [], 'p2': [], 'p3': []}
                        
                        # CẤU HÌNH TRANG A4 VÀ LỀ TRANG
                        sec = doc.sections[0]
                        sec.page_width = Cm(21.0)
                        sec.page_height = Cm(29.7)
                        sec.top_margin = Cm(1.5)
                        sec.bottom_margin = Cm(1.5)
                        sec.left_margin = Cm(2.0)
                        sec.right_margin = Cm(1.5)
                        
                        # 1. TIÊU ĐỀ HEADERS
                        table = doc.add_table(rows=1, cols=2)
                        table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        table.autofit = False
                        
                        cell_left, cell_right = table.rows[0].cells
                        cell_left.width = Cm(8.5)
                        cell_right.width = Cm(9.0)
                        
                        p_l = cell_left.paragraphs[0]
                        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_so = p_l.add_run(f"{so_gd.upper()}\n")
                        r_so.font.name = 'Times New Roman'; r_so.font.size = Pt(10)
                        
                        r_truong = p_l.add_run(f"{ten_truong.upper()}")
                        r_truong.bold = True
                        r_truong.font.name = 'Times New Roman'; r_truong.font.size = Pt(11)
                        
                        sec_notes = []
                        if len(phan1_goc) > 0: sec_notes.append(f"{len(phan1_goc)} TN")
                        if len(phan2_goc) > 0: sec_notes.append(f"{len(phan2_goc)} ĐS")
                        if len(phan3_goc) > 0: sec_notes.append(f"{len(phan3_goc)} TL Ngắn")
                        note_str = f"({', '.join(sec_notes)})" if sec_notes else ""
                        
                        p_r = cell_right.paragraphs[0]
                        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_kt = p_r.add_run(f"{ten_ky_thi.upper()}\n")
                        r_kt.bold = True
                        r_kt.font.name = 'Times New Roman'; r_kt.font.size = Pt(11)
                        
                        r_mon = p_r.add_run(f"Môn thi: {mon_thi.upper()}\n")
                        r_mon.bold = True
                        r_mon.font.name = 'Times New Roman'; r_mon.font.size = Pt(10.5)
                        
                        r_tg = p_r.add_run(f"Thời gian: {thoi_gian}\n")
                        r_tg.font.name = 'Times New Roman'; r_tg.font.size = Pt(10); r_tg.italic = True
                        
                        r_note = p_r.add_run(f"Đề có {note_str}")
                        r_note.font.name = 'Times New Roman'; r_note.font.size = Pt(9.5); r_note.italic = True
                        
                        # 2. ĐỊNH DẠNG HỌ TÊN, SỐ BÁO BÁO DẠNH, MÃ ĐỀ (3 DÒNG TÁCH BIỆT)
                        p_name = doc.add_paragraph()
                        p_name.paragraph_format.space_before = Pt(8)
                        p_name.paragraph_format.space_after = Pt(2)
                        r_name = p_name.add_run("Họ tên học sinh: ....................................................................................................")
                        r_name.font.name = 'Times New Roman'; r_name.font.size = Pt(12)

                        p_sbd = doc.add_paragraph()
                        p_sbd.paragraph_format.space_before = Pt(2)
                        p_sbd.paragraph_format.space_after = Pt(2)
                        r_sbd = p_sbd.add_run("Số báo danh: ........................................................................................................")
                        r_sbd.font.name = 'Times New Roman'; r_sbd.font.size = Pt(12)

                        p_made = doc.add_paragraph()
                        p_made.paragraph_format.space_before = Pt(2)
                        p_made.paragraph_format.space_after = Pt(10)
                        r_m1 = p_made.add_run("Mã đề: ")
                        r_m1.bold = True
                        r_m1.font.name = 'Times New Roman'; r_m1.font.size = Pt(12)
                        r_m2 = p_made.add_run(f"{ma_de}")
                        r_m2.bold = True
                        r_m2.font.name = 'Times New Roman'; r_m2.font.size = Pt(12)

                        nhan_p1_word = ["A.", "B.", "C.", "D."]
                        nhan_p1_excel = ["A", "B", "C", "D"]
                        nhan_p2_word = ["a)", "b)", "c)", "d)"]

                        # 3. TRỘN PHẦN I
                        if phan1_goc:
                            p_head1 = doc.add_paragraph()
                            p_head1.paragraph_format.space_before = Pt(6)
                            p_head1.paragraph_format.space_after = Pt(4)
                            r_h1 = p_head1.add_run(f"PHẦN I. Thí sinh trả lời từ câu 1 đến câu {len(phan1_goc)}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
                            r_h1.bold = True
                            r_h1.font.name = 'Times New Roman'; r_h1.font.size = Pt(12)

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
                                correct_label = "A"
                                for o_idx, opt in enumerate(opts):
                                    if opt["is_correct"]:
                                        correct_label = nhan_p1_excel[o_idx]
                                
                                Render_Phuong_An_Phan_1(doc, opts, nhan_p1_word)
                                ma_de_answers[ma_de]['p1'].append(correct_label)

                        # 4. TRỘN PHẦN II
                        if phan2_goc:
                            p_head2 = doc.add_paragraph()
                            p_head2.paragraph_format.space_before = Pt(10)
                            p_head2.paragraph_format.space_after = Pt(4)
                            r_h2 = p_head2.add_run(f"PHẦN II. Thí sinh trả lời từ câu 1 đến câu {len(phan2_goc)}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
                            r_h2.bold = True
                            r_h2.font.name = 'Times New Roman'; r_h2.font.size = Pt(12)

                            p2_shuffled = copy.deepcopy(phan2_goc)
                            random.shuffle(p2_shuffled)
                            for idx, q in enumerate(p2_shuffled):
                                updated_q = Cap_Nhat_Tien_To_Cau(q["q_elem"], f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q)
                                
                                subs = q["sub_options"]
                                random.shuffle(subs)
                                ans_chars = []
                                for s_idx, sub in enumerate(subs):
                                    prefix = f"{nhan_p2_word[s_idx]} "
                                    elem = Lay_Noi_Dung_Phuong_An_Clean(copy.deepcopy(sub["elem"]), prefix)
                                    append_elem(doc, elem)
                                    
                                    char = "Đ" if sub["is_correct"] == "Đúng" else "S"
                                    ans_chars.append(char)
                                    
                                ma_de_answers[ma_de]['p2'].append("".join(ans_chars))

                        # 5. TRỘN PHẦN III
                        if phan3_goc:
                            p_head3 = doc.add_paragraph()
                            p_head3.paragraph_format.space_before = Pt(10)
                            p_head3.paragraph_format.space_after = Pt(4)
                            r_h3 = p_head3.add_run(f"PHẦN III. Thí sinh trả lời từ câu 1 đến câu {len(phan3_goc)}.")
                            r_h3.bold = True
                            r_h3.font.name = 'Times New Roman'; r_h3.font.size = Pt(12)

                            p3_shuffled = copy.deepcopy(phan3_goc)
                            random.shuffle(p3_shuffled)
                            for idx, q in enumerate(p3_shuffled):
                                updated_q = Cap_Nhat_Tien_To_Cau(q["q_elem"], f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q)
                                ma_de_answers[ma_de]['p3'].append(str(q["correct"]).strip())

                        # 6. LỜI DẶN CHÂN TRANG
                        p_foot1 = doc.add_paragraph("---Hết---")
                        p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foot1.paragraph_format.space_before = Pt(12)
                        r_f1 = p_foot1.runs[0]
                        r_f1.font.name = 'Times New Roman'; r_f1.font.size = Pt(12)
                        
                        p_foot2 = doc.add_paragraph("Thí Sinh Không Được Sử Dụng Tài Liệu - Cán Bộ Coi Thi Không Giải Thích Gì Thêm.")
                        p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_f2 = p_foot2.runs[0]
                        r_f2.font.name = 'Times New Roman'; r_f2.font.size = Pt(10); r_f2.italic = True

                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"{ma_de}.docx", doc_io.getvalue())
                    
                    # --- XUẤT FILE EXCEL ĐÁP ÁN THEO ĐÚNG MẪU BẢNG MA TRẬN ---
                    wb_excel = openpyxl.Workbook()
                    ws_excel = wb_excel.active
                    ws_excel.title = "Sheet1"

                    header = ['Câu'] + [str(m) for m in danh_sach_ma_de]
                    ws_excel.append(header)

                    first_md = danh_sach_ma_de[0]
                    num_p1 = len(ma_de_answers[first_md]['p1'])
                    num_p2 = len(ma_de_answers[first_md]['p2'])
                    num_p3 = len(ma_de_answers[first_md]['p3'])

                    # Hàng Phần I (1 đến N câu)
                    for i in range(num_p1):
                        row = [i + 1]
                        for md in danh_sach_ma_de:
                            row.append(ma_de_answers[md]['p1'][i])
                        ws_excel.append(row)

                    # Hàng Phần II (Gộp 4 ý thành 1 ô chuỗi 4 ký tự: ĐSĐS...)
                    for i in range(num_p2):
                        row = [i + 1]
                        for md in danh_sach_ma_de:
                            row.append(ma_de_answers[md]['p2'][i])
                        ws_excel.append(row)

                    # Hàng Phần III (Định dạng General/Chuỗi)
                    for i in range(num_p3):
                        row = [i + 1]
                        for md in danh_sach_ma_de:
                            row.append(str(ma_de_answers[md]['p3'][i]))
                        ws_excel.append(row)

                    align_center = Alignment(horizontal='center', vertical='center')
                    font_header = Font(name='Calibri', size=11, bold=True)
                    font_cell = Font(name='Calibri', size=11)

                    for row in ws_excel.iter_rows(min_row=1, max_row=ws_excel.max_row, min_col=1, max_col=ws_excel.max_column):
                        for cell in row:
                            cell.alignment = align_center
                            cell.number_format = 'General'
                            if cell.row == 1:
                                cell.font = font_header
                            else:
                                cell.font = font_cell

                    excel_io = io.BytesIO()
                    wb_excel.save(excel_io)
                    zip_file.writestr("Bang_Dap_An_Cac_Ma_De.xlsx", excel_io.getvalue())

                st.success("🎉 Trộn đề thành công! Đã tạo file Excel đáp án ma trận chuẩn 100% mẫu quy định.")
                st.download_button(
                    label="📥 TẢI XUẤN TOÀN BỘ ĐỀ THI & FILE EXCEL ĐÁP ÁN (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Bo_De_Thi_Chuyen_Nghiep.zip",
                    mime="application/zip",
                    type="primary"
                )

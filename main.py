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

# 1. CẤU HÌNH TRANG WEB
st.set_page_config(page_title="Trộn Đề Kiểm tra Trắc Nghiệm (Vật Lí - Hóa Học)", page_icon="📝", layout="centered")

st.title("📝 Ứng Dụng Trộn Đề Kiểm Tra Trắc Nghiệm (Vật Lí - Hóa Học)")
st.caption("👨‍🏫 **Tác giả:** LQChau, BNDQuan - CLB STEM - THPT Thủ Đức")
st.write("Công cụ trộn đề kiểm tra theo cấu trúc Bộ GD&ĐT (Hỗ trợ tốt nhất cho công thức MathType/OMML, công thức Hóa học & Hình ảnh cấu tạo).")

st.divider()

# 2. TRÍCH XUẤT VĂN BẢN VÀ XỬ LÝ XML CÔNG THỨC & HÌNH ẢNH
def get_node_text(node):
    """Trích xuất toàn bộ chữ trong đoạn văn XML, bao gồm chữ thường (<w:t>) và công thức toán (<m:t>)"""
    if node.tag.endswith('}t'):
        return node.text or ""
    text = ""
    for child in node.iter():
        if child.tag.endswith('}t'):
            text += child.text or ""
    return text

def Thiet_Lap_Font_Times_New_Roman(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    rPr = style._element.get_or_add_rPr()
    rFonts = parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>')
    rPr.append(rFonts)

def Set_Paragraph_Justify_And_Zero_Indent(p):
    """Thiết lập căn đều 2 bên (Justify) và xóa 100% thụt lề (Left Indent, First Line Indent = 0)"""
    pPr = p._element.pPr
    if pPr is None:
        pPr = parse_xml(r'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        p._element.insert(0, pPr)
        
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is not None:
        pPr.remove(pStyle)
        
    numPr = pPr.find(qn('w:numPr'))
    if numPr is not None:
        pPr.remove(numPr)
        
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        pPr.remove(ind)
        
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.right_indent = Cm(0)
    
    jc = pPr.find(qn('w:jc'))
    if jc is not None:
        jc.set(qn('w:val'), 'both')
    else:
        pPr.append(parse_xml(r'<w:jc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="both"/>'))

def Copy_Element_With_Media(p_elem, doc_src, doc_target, rid_map):
    """Deep copy XML element và re-map chính xác các rId hình ảnh sang doc_target"""
    elem = copy.deepcopy(p_elem)
    for node in elem.iter():
        for attr_name in list(node.attrib.keys()):
            if attr_name.endswith('}embed') or attr_name.endswith('}id'):
                old_rid = node.attrib[attr_name]
                if old_rid in doc_src.part.rels:
                    if old_rid not in rid_map:
                        rel = doc_src.part.rels[old_rid]
                        if "image" in rel.reltype:
                            image_part = rel.target_part
                            new_rid = doc_target.part.relate_to(image_part, rel.reltype)
                            rid_map[old_rid] = new_rid
                        else:
                            try:
                                new_rid = doc_target.part.rels.add_relationship(
                                    rel.reltype,
                                    rel.target_ref if rel.is_external else rel.target_part,
                                    rel.is_external
                                )
                                rid_map[old_rid] = new_rid
                            except Exception:
                                pass
                    if old_rid in rid_map:
                        node.attrib[attr_name] = rid_map[old_rid]
    return elem

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

def Has_Auto_Numbering(p):
    pPr = p._element.pPr
    if pPr is not None and pPr.numPr is not None:
        return True
    return False

# 3. HÀM TÁCH CÁC PHƯƠNG ÁN NẰM CÙNG DÒNG (BỔ SUNG THUẬT TOÁN GỘP THẺ XML CHO ĐỀ HÓA)
def split_paragraph_options(p_elem):
    children = [child for child in p_elem if not child.tag.endswith('}pPr')]
    if not children:
        full_text = get_node_text(p_elem).strip()
        clean_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', full_text).strip()
        return [{"elem": p_elem, "text": clean_text, "is_under": False}]

    # Xây dựng chuỗi toàn văn và bản đồ ánh chỉ số ký tự sang chỉ số child XML
    full_str = ""
    char_to_child_idx = []
    for idx, child in enumerate(children):
        c_txt = get_node_text(child)
        if not c_txt and ("drawing" in child.xml or "graphic" in child.xml or "pict" in child.xml):
            c_txt = " [IMG] "
        for ch in c_txt:
            full_str += ch
            char_to_child_idx.append(idx)

    # Tìm vị trí bắt đầu các phương án A., B., C., D.
    pattern = re.compile(r'(?:^|[\s\t]+)([A-Da-d][\.\:\-\)\s])')
    matches = list(pattern.finditer(full_str))

    # Nếu không tìm thấy hoặc chỉ có 1 nhãn phương án ở đầu dòng
    if len(matches) <= 1:
        full_text = full_str.strip()
        clean_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', full_text).strip()
        is_under = False
        for r in p_elem.iter():
            if r.tag.endswith('}r'):
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    u = rPr.find(qn('w:u'))
                    if u is not None and u.get(qn('w:val')) != 'none':
                        is_under = True
                        break
        return [{"elem": p_elem, "text": clean_text, "is_under": is_under}]

    # Gom các child XML thành từng nhóm phương án tương ứng
    groups = []
    start_child_indices = []
    for m in matches:
        pos = m.start(1)
        c_idx = char_to_child_idx[pos] if pos < len(char_to_child_idx) else len(children) - 1
        if not start_child_indices or start_child_indices[-1] != c_idx:
            start_child_indices.append(c_idx)

    for i in range(len(start_child_indices)):
        c_start = start_child_indices[i]
        c_end = start_child_indices[i+1] if i + 1 < len(start_child_indices) else len(children)
        grp_children = children[c_start:c_end]
        
        new_p = copy.deepcopy(p_elem)
        for child in list(new_p):
            if not child.tag.endswith('}pPr'):
                new_p.remove(child)
                
        is_under = False
        grp_text = ""
        for child in grp_children:
            new_p.append(copy.deepcopy(child))
            grp_text += get_node_text(child)
            for r in child.iter():
                if r.tag.endswith('}r'):
                    rPr = r.find(qn('w:rPr'))
                    if rPr is not None:
                        u = rPr.find(qn('w:u'))
                        if u is not None and u.get(qn('w:val')) != 'none':
                            is_under = True

        clean_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', grp_text.strip()).strip()
        groups.append({
            "elem": new_p,
            "text": clean_text,
            "is_under": is_under
        })
    return groups

# 4. CẬP NHẬT TIỀN TỐ CÂU & PHƯƠNG ÁN
def Cap_Nhat_Tien_To_Cau(p_elem, doc_src, doc_target, rid_map, new_prefix):
    elem_copied = Copy_Element_With_Media(p_elem, doc_src, doc_target, rid_map)
    doc_temp = Document()
    append_elem(doc_temp, elem_copied)
    p = doc_temp.paragraphs[0]
    Set_Paragraph_Justify_And_Zero_Indent(p)
    
    full_text = get_node_text(p._element)
    match = re.match(r'^(Câu|Cau)?\s*\d+[\.\:\-\)\s]*', full_text, re.IGNORECASE)
    if match and match.group(0).strip():
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
    p._element.pPr.addnext(r0._element)
    return p._element

def Lay_Noi_Dung_Phuong_An_Clean(p_elem, doc_src, doc_target, rid_map, new_prefix):
    elem_copied = Copy_Element_With_Media(p_elem, doc_src, doc_target, rid_map)
    doc_temp = Document()
    append_elem(doc_temp, elem_copied)
    p = doc_temp.paragraphs[0]
    Set_Paragraph_Justify_And_Zero_Indent(p)
    
    for tab in p._element.xpath('.//w:tab'):
        tab.getparent().remove(tab)
        
    full_text = get_node_text(p._element)
    match = re.match(r'^[A-Da-d0-9][\.\:\-\)\s]*', full_text)
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
                
    if new_prefix:
        r0 = p.add_run(new_prefix)
        r0.bold = True
        r0.font.name = 'Times New Roman'
        r0.font.size = Pt(12)
        p._element.pPr.addnext(r0._element)
    return p._element

# 5. DÀN HÀNG ĐÁP ÁN PHẦN I CHUẨN THẨM MỸ
def Render_Phuong_An_Phan_1(doc, doc_src, rid_map, opts, nhan_p1):
    total_len = sum(len(opt["text"]) for opt in opts)
    max_len = max(len(opt["text"]) for opt in opts) if opts else 0
    has_math_or_img = any("oMath" in opt["elem"].xml or "drawing" in opt["elem"].xml or "graphic" in opt["elem"].xml for opt in opts)
    
    if total_len <= 28 and max_len <= 8 and len(opts) == 4 and not has_math_or_img:
        p_combined = Document().add_paragraph()
        Set_Paragraph_Justify_And_Zero_Indent(p_combined)
        p_format = p_combined.paragraph_format
        p_format.space_before = Pt(0)
        p_format.space_after = Pt(2)
        p_format.tab_stops.add_tab_stop(Cm(4.0))
        p_format.tab_stops.add_tab_stop(Cm(8.0))
        p_format.tab_stops.add_tab_stop(Cm(12.0))
        
        for idx, opt in enumerate(opts):
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(opt["elem"], doc_src, doc, rid_map, prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            for child in doc_tmp.paragraphs[0]._element:
                if not child.tag.endswith('}pPr'):
                    p_combined._element.append(copy.deepcopy(child))
            if idx < 3:
                r_t = p_combined.add_run("\t")
                r_t.font.name = 'Times New Roman'
                r_t.font.size = Pt(12)
                
        append_elem(doc, p_combined._element)
        
    elif total_len <= 100 and max_len <= 32 and len(opts) == 4 and not has_math_or_img:
        p_line1 = Document().add_paragraph()
        Set_Paragraph_Justify_And_Zero_Indent(p_line1)
        p_line1.paragraph_format.space_before = Pt(0)
        p_line1.paragraph_format.space_after = Pt(1)
        p_line1.paragraph_format.tab_stops.add_tab_stop(Cm(8.5))
        
        for idx in [0, 1]:
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(opts[idx]["elem"], doc_src, doc, rid_map, prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            for child in doc_tmp.paragraphs[0]._element:
                if not child.tag.endswith('}pPr'):
                    p_line1._element.append(copy.deepcopy(child))
            if idx == 0:
                r_t = p_line1.add_run("\t")
                r_t.font.name = 'Times New Roman'
                r_t.font.size = Pt(12)
        append_elem(doc, p_line1._element)
        
        p_line2 = Document().add_paragraph()
        Set_Paragraph_Justify_And_Zero_Indent(p_line2)
        p_line2.paragraph_format.space_before = Pt(0)
        p_line2.paragraph_format.space_after = Pt(2)
        p_line2.paragraph_format.tab_stops.add_tab_stop(Cm(8.5))
        
        for idx in [2, 3]:
            prefix = f"{nhan_p1[idx]} "
            elem = Lay_Noi_Dung_Phuong_An_Clean(opts[idx]["elem"], doc_src, doc, rid_map, prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            for child in doc_tmp.paragraphs[0]._element:
                if not child.tag.endswith('}pPr'):
                    p_line2._element.append(copy.deepcopy(child))
            if idx == 2:
                r_t = p_line2.add_run("\t")
                r_t.font.name = 'Times New Roman'
                r_t.font.size = Pt(12)
        append_elem(doc, p_line2._element)
        
    else:
        for idx, opt in enumerate(opts):
            prefix = f"{nhan_p1[idx]} " if idx < len(nhan_p1) else ""
            elem = Lay_Noi_Dung_Phuong_An_Clean(opt["elem"], doc_src, doc, rid_map, prefix)
            doc_tmp = Document()
            append_elem(doc_tmp, elem)
            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
            doc_tmp.paragraphs[0].paragraph_format.space_before = Pt(0)
            doc_tmp.paragraphs[0].paragraph_format.space_after = Pt(2)
            append_elem(doc, doc_tmp.paragraphs[0]._element)

# 6. ĐỌC VÀ XÁC ĐỊNH CÔNG THỨC & NỘI DUNG CÂU HỎI (CHUẨN HÓA CHO ĐỀ HÓA & VẬT LÍ)
def doc_de_goc(file_stream):
    doc = Document(file_stream)
    phan1, phan2, phan3 = [], [], []
    current_section = None
    current_item = None
    in_solution = False

    def is_sol_header(line):
        pattern = r'^(?:Lời\s*giải|Hướng\s*dẫn\s*giải|HDG|Giải\s*chi\s*tiết|Đáp\s*án|Đáp\s*số|ĐS|ĐA|Giải)\b[\:\-\.\s]*'
        return bool(re.match(pattern, line, re.IGNORECASE))

    def is_sec_header(line):
        return bool(re.search(r'(?:\[?\s*PHAN[_\s]*[123]\s*\]?|PHẦN\s*(I|II|III|1|2|3)\b)', line, re.IGNORECASE))

    for p in doc.paragraphs:
        full_line = get_node_text(p._element).strip()
        has_drawing = ("drawing" in p._element.xml or "graphic" in p._element.xml or "pict" in p._element.xml or "oMath" in p._element.xml)
        
        if not full_line and not has_drawing:
            continue

        if re.search(r'(?:\[?\s*PHAN[_\s]*3\s*\]?|PHẦN\s*(III|3)\b)', full_line, re.IGNORECASE):
            if current_item:
                if current_section == 1: phan1.append(current_item)
                elif current_section == 2: phan2.append(current_item)
                elif current_section == 3: phan3.append(current_item)
                current_item = None
            current_section = 3
            in_solution = False
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*2\s*\]?|PHẦN\s*(II|2)\b)', full_line, re.IGNORECASE):
            if current_item:
                if current_section == 1: phan1.append(current_item)
                elif current_section == 2: phan2.append(current_item)
                elif current_section == 3: phan3.append(current_item)
                current_item = None
            current_section = 2
            in_solution = False
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*1\s*\]?|PHẦN\s*(I|1)\b)', full_line, re.IGNORECASE):
            if current_item:
                if current_section == 1: phan1.append(current_item)
                elif current_section == 2: phan2.append(current_item)
                elif current_section == 3: phan3.append(current_item)
                current_item = None
            current_section = 1
            in_solution = False
            continue

        # Kiểm tra tiêu đề câu hỏi chính xác (tránh nhầm các ý (1), (2)... trong đề Hóa)
        is_explicit_q = bool(re.match(r'^(Câu|Cau)\s*\d+[\.\:\-\)\s]', full_line, re.IGNORECASE))
        is_opt_line = bool(re.match(r'^[A-Da-d][\.\:\-\)\s]', full_line))
        is_auto_num = Has_Auto_Numbering(p)
        
        is_question = False
        if not is_sec_header(full_line) and not is_sol_header(full_line):
            if is_explicit_q:
                is_question = True
            elif is_auto_num and not is_opt_line:
                is_question = True

        if is_question:
            in_solution = False

        if is_sol_header(full_line):
            in_solution = True

        if current_section == 1:
            if is_question:
                if current_item: phan1.append(current_item)
                current_item = {"q_elems": [p._element], "options": [], "correct_idx": None}
            elif current_item and not in_solution:
                if is_opt_line:
                    extracted_opts = split_paragraph_options(p._element)
                    for opt in extracted_opts:
                        opt_idx = len(current_item["options"])
                        current_item["options"].append({
                            "elem": opt["elem"],
                            "text": opt["text"]
                        })
                        if opt["is_under"]:
                            current_item["correct_idx"] = opt_idx
                elif len(current_item["options"]) == 0:
                    current_item["q_elems"].append(p._element)
            elif current_item and in_solution:
                if current_item["correct_idx"] is None:
                    m_ans = re.search(r'(?:Chọn|Đáp\s*án|ĐA)\s*([A-D])\b', full_line, re.IGNORECASE)
                    if m_ans:
                        ans_char = m_ans.group(1).upper()
                        current_item["correct_idx"] = ord(ans_char) - ord('A')

        elif current_section == 2:
            if is_question:
                if current_item: phan2.append(current_item)
                current_item = {"q_elems": [p._element], "sub_options": []}
            elif current_item and not in_solution:
                if is_opt_line:
                    clean_sub_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', full_line).strip()
                    is_under = Kiem_Tra_Gach_Chan(p)
                    current_item["sub_options"].append({
                        "elem": p._element,
                        "text": clean_sub_text,
                        "is_correct": "Đúng" if is_under else "Sai"
                    })
                elif len(current_item["sub_options"]) == 0:
                    current_item["q_elems"].append(p._element)

        elif current_section == 3:
            if is_question:
                if current_item: phan3.append(current_item)
                current_item = {"q_elems": [p._element], "correct": ""}
            elif current_item:
                if is_sol_header(full_line):
                    if re.search(r'(?:Đáp\s*(?:số|án)|Đ[SÁA]|DS|DA)\b', full_line, re.IGNORECASE):
                        ans_val = re.sub(r'^.*?(?:Đáp\s*(?:số|án)|Đ[SÁA]|DS|DA)\s*[\:\-\.]*\s*', '', full_line, flags=re.IGNORECASE).strip()
                        if ans_val:
                            current_item["correct"] = ans_val
                elif in_solution:
                    if not current_item["correct"] and full_line.strip() and not full_line.lower().startswith("lời giải"):
                        current_item["correct"] = full_line.strip()
                elif not in_solution:
                    current_item["q_elems"].append(p._element)

    if current_item:
        if current_section == 1: phan1.append(current_item)
        elif current_section == 2: phan2.append(current_item)
        elif current_section == 3: phan3.append(current_item)

    return doc, phan1, phan2, phan3

# 7. GIAO DIỆN STREAMLIT & XỬ LÝ TRỘN ĐỀ
with st.expander("⚙️ **THÔNG TIN TIÊU ĐỀ ĐỀ KIỂM TRA & CÁC MÃ ĐỀ**", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        so_gd = st.text_input("Sở Giáo dục & Đào tạo", value="SỞ GIÁO DỤC VÀ ĐÀO TẠO TP.HCM")
        ten_truong = st.text_input("Tên Trường", value="TRƯỜNG THPT THỦ ĐỨC")
        ten_ky_thi = st.text_input("Tên Kỳ Thi", value="KIỂM TRA ĐỊNH KÌ HỌC KÌ I")
    with col2:
        mon_thi = st.text_input("Môn thi / Khối", value="MÔN HÓA HỌC - KHỐI 12")
        thoi_gian = st.text_input("Thời gian làm bài", value="45 phút")
        ds_ma_de_input = st.text_input("Danh sách Mã Đề (cách nhau bởi dấu phẩy)", value="1021, 1022, 1023, 1024")

uploaded_file = st.file_uploader("Bước 1: Tải lên file Word đề gốc (.docx)", type=["docx"])

if uploaded_file is not None:
    st.success(f"✅ Đã nhận file: **{uploaded_file.name}**")

if st.button("🚀 Bước 2: Bắt Đầu Trộn Đề & Tạo File Đáp Án", type="primary"):
    if uploaded_file is None:
        st.warning("⚠️ Vui lòng chọn file Word đề gốc ở Bước 1 trước!")
    else:
        with st.spinner("Đang khởi tạo mã đề thi và xuất file Excel đáp án ma trận..."):
            doc_goc_obj, phan1_goc, phan2_goc, phan3_goc = doc_de_goc(uploaded_file)
            
            danh_sach_ma_de = [m.strip() for m in ds_ma_de_input.split(",") if m.strip()]
            if not danh_sach_ma_de:
                danh_sach_ma_de = ["1021", "1022", "1023", "1024"]
                
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
                        rid_map = {}
                        ma_de_answers[ma_de] = {'p1': [], 'p2': [], 'p3': []}
                        
                        sec = doc.sections[0]
                        sec.page_width = Cm(21.0)
                        sec.page_height = Cm(29.7)
                        sec.top_margin = Cm(1.5)
                        sec.bottom_margin = Cm(1.5)
                        sec.left_margin = Cm(2.0)
                        sec.right_margin = Cm(1.5)
                        
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

                        # TRỘN PHẦN I
                        if phan1_goc:
                            p_head1 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head1)
                            p_head1.paragraph_format.space_before = Pt(6)
                            p_head1.paragraph_format.space_after = Pt(4)
                            r_h1 = p_head1.add_run(f"PHẦN I. Thí sinh trả lời từ câu 1 đến câu {len(phan1_goc)}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
                            r_h1.bold = True
                            r_h1.font.name = 'Times New Roman'; r_h1.font.size = Pt(12)

                            p1_shuffled = copy.deepcopy(phan1_goc)
                            random.shuffle(p1_shuffled)
                            for idx, q in enumerate(p1_shuffled):
                                updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q0)
                                
                                for sub_elem in q["q_elems"][1:]:
                                    sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                    doc_tmp = Document()
                                    append_elem(doc_tmp, sub_copied)
                                    Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                    append_elem(doc, doc_tmp.paragraphs[0]._element)
                                
                                opts = q["options"]
                                correct_idx = q["correct_idx"]
                                for o_i, opt in enumerate(opts):
                                    opt["is_correct"] = (o_i == correct_idx)
                                
                                random.shuffle(opts)
                                correct_label = "A"
                                for o_idx, opt in enumerate(opts):
                                    if opt["is_correct"]:
                                        correct_label = nhan_p1_excel[o_idx] if o_idx < len(nhan_p1_excel) else "A"
                                
                                Render_Phuong_An_Phan_1(doc, doc_goc_obj, rid_map, opts, nhan_p1_word)
                                ma_de_answers[ma_de]['p1'].append(correct_label)

                        # TRỘN PHẦN II
                        if phan2_goc:
                            p_head2 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head2)
                            p_head2.paragraph_format.space_before = Pt(10)
                            p_head2.paragraph_format.space_after = Pt(4)
                            r_h2 = p_head2.add_run(f"PHẦN II. Thí sinh trả lời từ câu 1 đến câu {len(phan2_goc)}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
                            r_h2.bold = True
                            r_h2.font.name = 'Times New Roman'; r_h2.font.size = Pt(12)

                            p2_shuffled = copy.deepcopy(phan2_goc)
                            random.shuffle(p2_shuffled)
                            for idx, q in enumerate(p2_shuffled):
                                updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q0)
                                
                                for sub_elem in q["q_elems"][1:]:
                                    sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                    doc_tmp = Document()
                                    append_elem(doc_tmp, sub_copied)
                                    Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                    append_elem(doc, doc_tmp.paragraphs[0]._element)
                                
                                subs = q["sub_options"]
                                random.shuffle(subs)
                                ans_chars = []
                                for s_idx, sub in enumerate(subs):
                                    prefix = f"{nhan_p2_word[s_idx]} " if s_idx < len(nhan_p2_word) else ""
                                    elem = Lay_Noi_Dung_Phuong_An_Clean(sub["elem"], doc_goc_obj, doc, rid_map, prefix)
                                    append_elem(doc, elem)
                                    
                                    char = "Đ" if sub["is_correct"] == "Đúng" else "S"
                                    ans_chars.append(char)
                                    
                                ma_de_answers[ma_de]['p2'].append("".join(ans_chars))

                        # TRỘN PHẦN III
                        if phan3_goc:
                            p_head3 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head3)
                            p_head3.paragraph_format.space_before = Pt(10)
                            p_head3.paragraph_format.space_after = Pt(4)
                            r_h3 = p_head3.add_run(f"PHẦN III. Thí sinh trả lời từ câu 1 đến câu {len(phan3_goc)}.")
                            r_h3.bold = True
                            r_h3.font.name = 'Times New Roman'; r_h3.font.size = Pt(12)

                            p3_shuffled = copy.deepcopy(phan3_goc)
                            random.shuffle(p3_shuffled)
                            for idx, q in enumerate(p3_shuffled):
                                updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {idx + 1}: ")
                                append_elem(doc, updated_q0)
                                
                                for sub_elem in q["q_elems"][1:]:
                                    sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                    doc_tmp = Document()
                                    append_elem(doc_tmp, sub_copied)
                                    Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                    append_elem(doc, doc_tmp.paragraphs[0]._element)
                                    
                                ma_de_answers[ma_de]['p3'].append(str(q["correct"]).strip())

                        # FOOTER CHÂN TRANG
                        p_foot1 = doc.add_paragraph()
                        p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foot1.paragraph_format.space_before = Pt(12)
                        p_foot1.paragraph_format.space_after = Pt(2)
                        r_f1 = p_foot1.add_run("---HẾT---")
                        r_f1.bold = True
                        r_f1.font.name = 'Times New Roman'
                        r_f1.font.size = Pt(12)
                        
                        p_foot2 = doc.add_paragraph("Thí Sinh Không Được Sử Dụng Tài Liệu - Cán Bộ Coi Thi Không Giải Thích Gì Thêm.")
                        p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_f2 = p_foot2.runs[0]
                        r_f2.font.name = 'Times New Roman'; r_f2.font.size = Pt(10); r_f2.italic = True

                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"{ma_de}.docx", doc_io.getvalue())
                    
                    # XUẤT FILE EXCEL ĐÁP ÁN MA TRẬN
                    wb_excel = openpyxl.Workbook()
                    ws_excel = wb_excel.active
                    ws_excel.title = "Sheet1"

                    header = ['Câu']
                    for m in danh_sach_ma_de:
                        header.append(int(m) if str(m).isdigit() else str(m))
                    ws_excel.append(header)

                    first_md = danh_sach_ma_de[0]
                    num_p1 = len(ma_de_answers[first_md]['p1'])
                    num_p2 = len(ma_de_answers[first_md]['p2'])
                    num_p3 = len(ma_de_answers[first_md]['p3'])

                    for i in range(num_p1):
                        row = [i + 1]
                        for md in danh_sach_ma_de:
                            row.append(ma_de_answers[md]['p1'][i])
                        ws_excel.append(row)

                    for i in range(num_p2):
                        row = [i + 1]
                        for md in danh_sach_ma_de:
                            row.append(ma_de_answers[md]['p2'][i])
                        ws_excel.append(row)

                    for i in range(num_p3):
                        row = [i + 1]
                        for md in danh_sach_ma_de:
                            val = str(ma_de_answers[md]['p3'][i]).strip()
                            if val.isdigit():
                                val = int(val)
                            row.append(val)
                        ws_excel.append(row)

                    align_center = Alignment(horizontal='center', vertical='center')
                    font_header = Font(name='Calibri', size=11, bold=True)
                    font_cell = Font(name='Calibri', size=11, bold=False)

                    for row in ws_excel.iter_rows(min_row=1, max_row=ws_excel.max_row, min_col=1, max_col=ws_excel.max_column):
                        for cell in row:
                            cell.alignment = align_center
                            cell.number_format = 'General'
                            if cell.row == 1:
                                cell.font = font_header
                            else:
                                cell.font = font_cell

                    ws_excel.column_dimensions['A'].width = 9.0

                    excel_io = io.BytesIO()
                    wb_excel.save(excel_io)
                    zip_file.writestr("Dap_an_cac_Ma_de.xlsx", excel_io.getvalue())

                st.success("🎉 Trộn đề thành công! Đã xử lý trọn vẹn 100% câu hỏi Hóa học/Vật lí, hình ảnh công thức cấu tạo & đáp án!")
                st.download_button(
                    label="📥 TẢI XUẤT TOÀN BỘ ĐỀ THI & FILE EXCEL ĐÁP ÁN (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="DE_TRON_HOA_VATLI.zip",
                    mime="application/zip",
                    type="primary"
                )

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
st.set_page_config(page_title="Phần mềm Trộn Đề Kiểm tra môn Vật Lí", page_icon="📝", layout="centered")

st.title("📝 Ứng Dụng Trộn Đề Kiểm Tra Môn Vật Lí Online")
st.caption("👨‍🏫 **Tác giả:** LQChau, BNDQuan - CLB STEM - THPT Thủ Đức")
st.write("Công cụ trộn đề kiểm tra môn Vật lí chuẩn Bộ GD&ĐT (Bảo toàn 100% Hình ảnh, Đồ thị, Bảng biểu, MathType & Font Times New Roman 12pt).")

st.divider()

# 2. HÀM THIẾT LẬP FONT VÀ XỬ LÝ COPY HÌNH ẢNH CHUẨN XÁC MÃ RID
def Thiet_Lap_Font_Times_New_Roman(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    rPr = style._element.get_or_add_rPr()
    rFonts = parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>')
    rPr.append(rFonts)

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

# 3. CẬP NHẬT TIỀN TỐ CÂU & PHƯƠNG ÁN (FONT TIMES NEW ROMAN 12PT)
def Cap_Nhat_Tien_To_Cau(p_elem, doc_src, doc_target, rid_map, new_prefix):
    elem_copied = Copy_Element_With_Media(p_elem, doc_src, doc_target, rid_map)
    doc_temp = Document()
    append_elem(doc_temp, elem_copied)
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

def Lay_Noi_Dung_Phuong_An_Clean(p_elem, doc_src, doc_target, rid_map, new_prefix):
    elem_copied = Copy_Element_With_Media(p_elem, doc_src, doc_target, rid_map)
    doc_temp = Document()
    append_elem(doc_temp, elem_copied)
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

# 4. DÀN HÀNG ĐÁP ÁN PHẦN I BẰNG TAB
def Render_Phuong_An_Phan_1(doc, doc_src, rid_map, opts, nhan_p1):
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
            elem = Lay_Noi_Dung_Phuong_An_Clean(opt["elem"], doc_src, doc, rid_map, prefix)
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
            elem = Lay_Noi_Dung_Phuong_An_Clean(opts[idx]["elem"], doc_src, doc, rid_map, prefix)
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
            elem = Lay_Noi_Dung_Phuong_An_Clean(opts[idx]["elem"], doc_src, doc, rid_map, prefix)
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
            elem = Lay_Noi_Dung_Phuong_An_Clean(opt["elem"], doc_src, doc, rid_map, prefix)
            append_elem(doc, elem)

# 5. ĐỌC VÀ BẢO TOÀN TRỌN VẸN NỘI DUNG THÂN CÂU HỎI (BAO GỒM CẢ HÌNH ẢNH)
def doc_de_goc(file_stream):
    doc = Document(file_stream)
    phan1, phan2, phan3 = [], [], []
    current_section = None
    current_item = None

    for p in doc.paragraphs:
        line = p.text.strip()
        has_drawing = ("drawing" in p._element.xml or "graphic" in p._element.xml or "pict" in p._element.xml)
        
        if not line and not has_drawing:
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
                current_item = {"q_elems": [p._element], "options": [], "correct_idx": None}
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
            elif current_item:
                if len(current_item["options"]) == 0:
                    current_item["q_elems"].append(p._element)

        elif current_section == 2:
            if is_question:
                if current_item: phan2.append(current_item)
                current_item = {"q_elems": [p._element], "sub_options": []}
            elif re.match(r'^[A-Da-d][\.\:\-\)\s]', line) and current_item:
                clean_sub_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', line).strip()
                is_under = Kiem_Tra_Gach_Chan(p)
                current_item["sub_options"].append({
                    "elem": p._element,
                    "text": clean_sub_text,
                    "is_correct": "Đúng" if is_under else "Sai"
                })
            elif current_item:
                if len(current_item["sub_options"]) == 0:
                    current_item["q_elems"].append(p._element)

        elif current_section == 3:
            if is_question:
                if current_item: phan3.append(current_item)
                current_item = {"q_elems": [p._element], "correct": ""}
            elif current_item:
                if re.search(r'(?:Đáp\s*(?:số|án)|Đ[SÁA]|DS|DA|Đáp)\b', line, re.IGNORECASE):
                    current_item["correct"] = re.sub(r'^.*?(?:Đáp\s*(?:số|án)|Đ[SÁA]|DS|DA|Đáp)\s*[\:\-\.]*\s*', '', line, flags=re.IGNORECASE).strip()
                elif not line.lower().startswith("lời giải") and not current_item["correct"]:
                    current_item["correct"] = line.strip()
                else:
                    current_item["q_elems"].append(p._element)

    if current_item:
        if current_section == 1: phan1.append(current_item)
        elif current_section == 2: phan2.append(current_item)
        elif current_section == 3: phan3.append(current_item)

    return doc, phan1, phan2, phan3

# 6. KHU VỰC THIẾT LẬP THÔNG TIN (UI STREAMLIT)
with st.expander("⚙️ **THÔNG TIN TIÊU ĐỀ ĐỀ KIỂM TRA & CÁC MÃ ĐỀ**", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        so_gd

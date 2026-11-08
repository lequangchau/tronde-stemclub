import streamlit as st
import os
import re
import random
import zipfile
import io
import copy
import urllib.request
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Alignment, Font
import pandas as pd
from lxml import etree

# 1. CẤU HÌNH TRANG WEB
st.set_page_config(page_title="LQChau - Trộn Đề Kiểm tra môn Vật Lí", page_icon="📝", layout="centered")

st.title("📝 Dự án CLB STEM: Ứng Dụng Trộn Đề Kiểm Tra Môn Vật Lí")
st.caption("👨‍🏫 **Tác giả:** LQChau, BNDQuan - CLB STEM - THPT Thủ Đức")
st.write("Công cụ trộn đề kiểm tra môn Vật lí theo cấu trúc của Bộ GD&ĐT (Hỗ trợ Câu hỏi chùm & Phần Tự luận)")

st.divider()

# CONFIG ĐƯỜNG DẪN FILE MẪU & LINK LƯU TRỮ ĐÁM MÂY (GOOGLE DRIVE CLOUD STORAGE)
FILE_MAU_TEN = "TX_B9_MIX.docx"  # Đặt file TX_B9_MIX.docx cùng thư mục với main.py
CLOUD_STORAGE_URL = "https://docs.google.com/document/d/1kNN3qXByIqA3pE9rBpat0JfY6v5iBGWi/edit?usp=sharing&ouid=109853281143982423134&rtpof=true&sd=true"

@st.cache_data
def load_file_mau_bytes():
    """Tải nội dung binary của file đề mẫu TX_B9_MIX.docx từ cục bộ hoặc từ Google Drive Cloud Storage"""
    if os.path.exists(FILE_MAU_TEN):
        try:
            with open(FILE_MAU_TEN, "rb") as f:
                return f.read()
        except Exception:
            pass
            
    if CLOUD_STORAGE_URL and CLOUD_STORAGE_URL.startswith("http"):
        try:
            download_url = CLOUD_STORAGE_URL
            match = re.search(r'google\.com/document/d/([^/]+)', CLOUD_STORAGE_URL)
            if match:
                doc_id = match.group(1)
                download_url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"

            req = urllib.request.Request(download_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                return response.read()
        except Exception:
            pass
            
    return None

# KHỐI GIAO DIỆN TẢI FILE MẪU CHUẨN (TX_B9_MIX.docx)
with st.container():
    col_mau1, col_mau2 = st.columns([1.8, 1], vertical_alignment="center")
    with col_mau1:
        st.markdown(f"📄 **File đề mẫu chuẩn:** Tải file **`{FILE_MAU_TEN}`** để xem định dạng")
    with col_mau2:
        file_mau_data = load_file_mau_bytes()
        if file_mau_data:
            st.download_button(
                label="📥 File Mẫu (.docx)",
                data=file_mau_data,
                file_name=FILE_MAU_TEN,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        elif CLOUD_STORAGE_URL:
            st.link_button(
                label="☁️ Tải Từ Cloud",
                url=CLOUD_STORAGE_URL,
                use_container_width=True
            )
        else:
            st.warning(f"⚠️ Đặt file `{FILE_MAU_TEN}` cùng thư mục code!")

st.divider()

# 2. HÀM XỬ LÝ AN TOÀN XML CÔNG THỨC & HÌNH ẢNH
def get_elem_xml(node):
    """Lấy chuỗi XML an toàn từ bất kỳ phần tử XML/lxml nào"""
    if hasattr(node, 'xml'):
        return node.xml
    try:
        return etree.tostring(node, encoding='unicode')
    except Exception:
        return ""

def get_node_text(node):
    """Trích xuất toàn bộ chữ trong đoạn văn XML (bao gồm chữ thường <w:t>, công thức toán <m:t> và tab <w:tab/>)"""
    if node.tag.endswith('}t'):
        return node.text or ""
    elif node.tag.endswith('}tab'):
        return "\t"
    text = ""
    for child in node.iter():
        if child.tag.endswith('}t'):
            text += child.text or ""
        elif child.tag.endswith('}tab'):
            text += "\t"
    return text

def clean_mix_tags_from_elem(p_elem):
    """Xóa triệt để các thẻ [MIX...], [MIX FIXED...], [END MIX...], [CHÙM...], [/CHÙM...] khỏi XML đoạn văn"""
    pattern = re.compile(
        r'\[\/?\s*(?:END\s+MIX|ENDMIX|MIX(?:\s*FIXED|_FIXED)?|FIXED\s*MIX|FIXED_MIX|CHÙM(?:\s*FIXED|_FIXED)?|CHUM(?:\s*FIXED|_FIXED)?|CỤM(?:\s*FIXED|_FIXED)?|CUM(?:\s*FIXED|_FIXED)?)(?:[_\s]*\d+)?\s*\]',
        re.IGNORECASE
    )
    t_nodes = [node for node in p_elem.iter() if node.tag.endswith('}t') and node.text]
    if not t_nodes:
        return
        
    full_text = "".join(node.text for node in t_nodes)
    matches = list(pattern.finditer(full_text))
    if not matches:
        return
        
    char_map = []
    for node in t_nodes:
        for idx in range(len(node.text)):
            char_map.append((node, idx))
            
    to_delete = set()
    for m in matches:
        for pos in range(m.start(), m.end()):
            to_delete.add(pos)
            
    node_text_map = {node: [] for node in t_nodes}
    for pos, (node, char_idx) in enumerate(char_map):
        if pos not in to_delete:
            node_text_map[node].append(node.text[char_idx])
            
    for node in t_nodes:
        node.text = "".join(node_text_map[node])

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
    clean_mix_tags_from_elem(elem)
    for node in elem.iter():
        for attr_name in list(node.attrib.keys()):
            if attr_name.endswith('}embed') or attr_name.endswith('}id'):
                old_rid = node.attrib[attr_name]
                if old_rid in doc_src.part.rels:
                    if old_rid not in rid_map:
                        rel = doc_src.part.rels[old_rid]
                        try:
                            if "image" in rel.reltype:
                                new_rid = doc_target.part.relate_to(rel.target_part, rel.reltype)
                                rid_map[old_rid] = str(new_rid)
                            elif rel.is_external:
                                new_rid = doc_target.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
                                rid_map[old_rid] = str(new_rid)
                            else:
                                new_rid = doc_target.part.relate_to(rel.target_part, rel.reltype)
                                rid_map[old_rid] = str(new_rid)
                        except Exception:
                            pass
                    if old_rid in rid_map and rid_map[old_rid] and rid_map[old_rid] != 'False':
                        node.attrib[attr_name] = rid_map[old_rid]
    return elem

def append_elem(doc, elem):
    body = doc._element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(elem)
    else:
        body.append(elem)

def Kiem_Tra_Gach_Chan_Elem(p_elem):
    if not p_elem.tag.endswith('p'):
        return False
    for r in p_elem.findall(qn('w:r')):
        t_text = "".join(t.text or "" for t in r.findall(qn('w:t')))
        if t_text.strip():
            rPr = r.find(qn('w:rPr'))
            if rPr is not None:
                u = rPr.find(qn('w:u'))
                if u is not None and u.get(qn('w:val')) != 'none':
                    return True
    return False

def Has_Auto_Numbering_Elem(p_elem):
    if not p_elem.tag.endswith('p'):
        return False
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is not None and pPr.find(qn('w:numPr')) is not None:
        return True
    return False

# 3. HÀM TÁCH CÁC PHƯƠNG ÁN NẰM CÙNG DÒNG VĂN BẢN
def split_paragraph_options(p_elem):
    children = [child for child in p_elem if not child.tag.endswith('}pPr')]
    if not children:
        full_text = get_node_text(p_elem).strip()
        clean_text = re.sub(r'^[\s\t\xa0]*[A-Da-d][\.\:\-\)\s\xa0]*', '', full_text).strip()
        return [{"elem": p_elem, "text": clean_text, "is_under": False}]

    full_str = ""
    char_to_child_idx = []
    for idx, child in enumerate(children):
        c_txt = get_node_text(child)
        child_xml = get_elem_xml(child)
        if not c_txt and ("drawing" in child_xml or "graphic" in child_xml or "pict" in child_xml or "object" in child_xml):
            c_txt = " [IMG] "
        for ch in c_txt:
            full_str += ch
            char_to_child_idx.append(idx)

    cand_pattern = re.compile(r'(?:^|[\s\t\xa0]+)([A-Da-d][\.\:\-\)])')
    all_cands = list(cand_pattern.finditer(full_str))

    if not all_cands:
        full_text = full_str.strip()
        clean_text = re.sub(r'^[\s\t\xa0]*[A-Da-d][\.\:\-\)\s\xa0]*', '', full_text).strip()
        return [{"elem": p_elem, "text": clean_text, "is_under": False}]

    valid_matches = []
    expected_char_codes = None
    
    for m in all_cands:
        lbl = m.group(1)[0]
        lbl_code = ord(lbl.upper())
        
        if not valid_matches:
            if full_str[:m.start(1)].strip() == "":
                valid_matches.append(m)
                expected_char_codes = [lbl_code + 1]
        else:
            if lbl_code in expected_char_codes:
                valid_matches.append(m)
                expected_char_codes = [lbl_code + 1]

    if len(valid_matches) <= 1:
        full_text = full_str.strip()
        clean_text = re.sub(r'^[\s\t\xa0]*[A-Da-d][\.\:\-\)\s\xa0]*', '', full_text).strip()
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

    groups = []
    start_child_indices = []
    for m in valid_matches:
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

        clean_text = re.sub(r'^[\s\t\xa0]*[A-Da-d][\.\:\-\)\s\xa0]*', '', grp_text.strip()).strip()
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
    
    for tab in p._element.xpath('.//w:tab'):
        tab.getparent().remove(tab)
        
    full_text = get_node_text(p._element)
    match = re.match(r'^[\s\t\xa0]*(?:(?:Câu|Cau)\s*\d+|\d+)[\.\:\-\)][\s\t\xa0]*', full_text, re.IGNORECASE)
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
    match = re.match(r'^[\s\t\xa0]*[A-Da-d][\.\:\-\)][\s\t\xa0]*', full_text)
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
    has_math = any("oMath" in get_elem_xml(opt["elem"]) for opt in opts)
    
    if total_len <= 28 and max_len <= 8 and len(opts) == 4 and not has_math:
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
        
    elif total_len <= 100 and max_len <= 32 and len(opts) == 4:
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

def Render_Cau_Dan_Chum(doc, start_q, end_q):
    """Tự động tạo câu dẫn mở đầu cho khối câu hỏi chùm với định dạng chuẩn"""
    p = doc.add_paragraph()
    Set_Paragraph_Justify_And_Zero_Indent(p)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    
    lead_text = f"Thông tin dưới đây được sử dụng từ câu {start_q} đến câu {end_q}:"
    r = p.add_run(lead_text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.bold = True
    r.italic = True

def Tao_Header_Trang_In(doc, so_gd, ten_truong, ten_ky_thi, mon_thi, thoi_gian, note_str, ma_de):
    """Tạo bảng khung tiêu đề chuẩn cho file Word đề thi và file Word đáp án"""
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
    
    if note_str:
        r_note = p_r.add_run(f"Đề có {note_str}")
        r_note.font.name = 'Times New Roman'; r_note.font.size = Pt(9.5); r_note.italic = True

    if ma_de:
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

# 6. ĐỌC VÀ XÁC ĐỊNH CẤU TRÚC ĐỀ GỐC (PHẦN I, II, III, IV)
def doc_de_goc(file_stream):
    doc = Document(file_stream)
    phan1_blocks, phan2_blocks, phan3_blocks, phan4_blocks = [], [], [], []
    current_section = None
    current_item = None
    current_cluster = None
    in_cluster = False
    in_solution = False

    def is_sol_header(line):
        pattern = r'^(?:Lời\s*giải|Hướng\s*dẫn\s*giải|HDG|Giải\s*chi\s*tiết|Đáp\s*án|Đáp\s*số|ĐS|ĐA|Giải)\b[\:\-\.\s]*'
        return bool(re.match(pattern, line, re.IGNORECASE))

    def is_sec_header(line):
        return bool(re.search(r'(?:\[?\s*PHAN[_\s]*[1234]\s*\]?|PHẦN\s*(I|II|III|IV|1|2|3|4)\b)', line, re.IGNORECASE))

    def is_mix_end(line):
        return bool(re.search(r'\[\/?\s*(?:END\s+MIX|ENDMIX|\/MIX|\/CHÙM|\/CHUM|\/CỤM|\/CUM)(?:[_\s]*\d+)?\s*\]', line, re.IGNORECASE))

    def is_mix_start(line):
        return bool(re.search(r'\[\s*(?:MIX|CHÙM|CHUM|CỤM|CUM)', line, re.IGNORECASE)) and not is_mix_end(line)

    def flush_item_to_current_block():
        nonlocal current_item, current_cluster, in_cluster, current_section
        if not current_item:
            return
        if current_section == 1:
            target_blocks = phan1_blocks
        elif current_section == 2:
            target_blocks = phan2_blocks
        elif current_section == 3:
            target_blocks = phan3_blocks
        else:
            target_blocks = phan4_blocks

        if in_cluster and current_cluster:
            current_cluster["items"].append(current_item)
        else:
            target_blocks.append({"is_cluster": False, "item": current_item})
        current_item = None

    def flush_cluster_to_blocks():
        nonlocal current_cluster, in_cluster, current_section
        flush_item_to_current_block()
        if in_cluster and current_cluster:
            if current_section == 1:
                target_blocks = phan1_blocks
            elif current_section == 2:
                target_blocks = phan2_blocks
            elif current_section == 3:
                target_blocks = phan3_blocks
            else:
                target_blocks = phan4_blocks

            if len(current_cluster["items"]) > 0:
                target_blocks.append(current_cluster)
            current_cluster = None
            in_cluster = False

    # LỌC AN TOÀN CHỈ CÁC PHẦN TỬ NỘI DUNG (ĐOẠN VĂN <w:p> & BẢNG BIỂU <w:tbl>)
    for elem in doc._element.body:
        if not (elem.tag.endswith('p') or elem.tag.endswith('tbl')):
            continue

        full_line = get_node_text(elem).strip()
        elem_xml = get_elem_xml(elem)
        has_drawing = ("drawing" in elem_xml or "graphic" in elem_xml or "pict" in elem_xml or "oMath" in elem_xml)
        
        if not full_line and not has_drawing:
            continue

        # Kiểm tra chuyển phần (Phần IV, III, II, I)
        if re.search(r'(?:\[?\s*PHAN[_\s]*4\s*\]?|PHẦN\s*(IV|4)\b)', full_line, re.IGNORECASE):
            flush_cluster_to_blocks()
            current_section = 4
            in_solution = False
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*3\s*\]?|PHẦN\s*(III|3)\b)', full_line, re.IGNORECASE):
            flush_cluster_to_blocks()
            current_section = 3
            in_solution = False
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*2\s*\]?|PHẦN\s*(II|2)\b)', full_line, re.IGNORECASE):
            flush_cluster_to_blocks()
            current_section = 2
            in_solution = False
            continue
        elif re.search(r'(?:\[?\s*PHAN[_\s]*1\s*\]?|PHẦN\s*(I|1)\b)', full_line, re.IGNORECASE):
            flush_cluster_to_blocks()
            current_section = 1
            in_solution = False
            continue

        # Kiểm tra thẻ [END MIX]
        if is_mix_end(full_line):
            flush_cluster_to_blocks()
            continue

        # Kiểm tra thẻ [MIX] hoặc [MIX FIXED]
        if is_mix_start(full_line):
            flush_cluster_to_blocks()
            in_cluster = True
            is_fixed = bool(re.search(r'FIXED', full_line, re.IGNORECASE))
            current_cluster = {
                "is_cluster": True,
                "is_fixed": is_fixed,
                "intro_elems": [],
                "items": []
            }
            clean_text = re.sub(r'\[\/?\s*(?:END\s+MIX|ENDMIX|MIX(?:\s*FIXED|_FIXED)?|FIXED\s*MIX|CHÙM(?:\s*FIXED|_FIXED)?|CHUM|CỤM|CUM)(?:[_\s]*\d+)?\s*\]', '', full_line, flags=re.IGNORECASE).strip()
            if not clean_text and not has_drawing:
                continue

        is_explicit_q = bool(re.match(r'^(Câu|Cau)\s*\d+[\.\:\-\)\s]', full_line, re.IGNORECASE)) or bool(re.match(r'^\d+[\.\:\-\)]', full_line))
        is_auto_num = Has_Auto_Numbering_Elem(elem)
        is_opt_line = bool(re.match(r'^[\s\t\xa0]*[A-Da-d][\.\:\-\)]', full_line))
        
        is_question = False
        if not is_sec_header(full_line) and not is_sol_header(full_line) and not is_opt_line:
            if is_explicit_q:
                is_question = True
            elif is_auto_num:
                if current_section in [1, 2]:
                    if current_item is None or len(current_item.get("options", [])) > 0 or len(current_item.get("sub_options", [])) > 0:
                        is_question = True
                elif current_section in [3, 4]:
                    if current_item is None or len(current_item.get("sol_elems", [])) > 0 or current_item.get("correct"):
                        is_question = True

        if is_question:
            in_solution = False
            flush_item_to_current_block()

        if is_sol_header(full_line):
            in_solution = True

        # Đọc dữ liệu theo từng Phần
        if current_section == 1:
            if is_question:
                current_item = {"q_elems": [elem], "options": [], "correct_idx": None}
            elif current_item and not in_solution:
                if is_opt_line:
                    extracted_opts = split_paragraph_options(elem)
                    for opt in extracted_opts:
                        opt_idx = len(current_item["options"])
                        current_item["options"].append({
                            "elem": opt["elem"],
                            "text": opt["text"]
                        })
                        if opt["is_under"]:
                            current_item["correct_idx"] = opt_idx
                elif len(current_item["options"]) == 0:
                    current_item["q_elems"].append(elem)
            elif current_item and in_solution:
                if current_item["correct_idx"] is None:
                    m_ans = re.search(r'(?:Chọn|Đáp\s*án|ĐA)\s*([A-D])\b', full_line, re.IGNORECASE)
                    if m_ans:
                        ans_char = m_ans.group(1).upper()
                        current_item["correct_idx"] = ord(ans_char) - ord('A')
            elif current_item is None and in_cluster and current_cluster:
                current_cluster["intro_elems"].append(elem)

        elif current_section == 2:
            if is_question:
                current_item = {"q_elems": [elem], "sub_options": []}
            elif current_item and not in_solution:
                if is_opt_line:
                    clean_sub_text = re.sub(r'^[A-Da-d][\.\:\-\)\s]*', '', full_line).strip()
                    is_under = Kiem_Tra_Gach_Chan_Elem(elem)
                    current_item["sub_options"].append({
                        "elem": elem,
                        "text": clean_sub_text,
                        "is_correct": "Đúng" if is_under else "Sai"
                    })
                elif len(current_item["sub_options"]) == 0:
                    current_item["q_elems"].append(elem)
            elif current_item is None and in_cluster and current_cluster:
                current_cluster["intro_elems"].append(elem)

        elif current_section == 3:
            if is_question:
                current_item = {"q_elems": [elem], "correct": ""}
            elif current_item:
                if is_sol_header(full_line):
                    m_ans = re.search(r'(?:Đáp\s*(?:số|án)|Lời\s*giải|HDG|Hướng\s*dẫn\s*giải|Giải|Đ[SÁA]|DS|DA)\s*[\:\-\.]*\s*(.*)$', full_line, re.IGNORECASE)
                    if m_ans and m_ans.group(1).strip():
                        current_item["correct"] = m_ans.group(1).strip()
                elif in_solution:
                    if not current_item["correct"] and full_line.strip() and not full_line.lower().startswith("lời giải"):
                        current_item["correct"] = full_line.strip()
                elif not in_solution:
                    current_item["q_elems"].append(elem)
            elif current_item is None and in_cluster and current_cluster:
                current_cluster["intro_elems"].append(elem)

        elif current_section == 4:
            if is_question:
                current_item = {"q_elems": [elem], "sol_elems": []}
            elif current_item:
                if is_sol_header(full_line):
                    in_solution = True
                    current_item["sol_elems"].append(elem)
                elif in_solution:
                    current_item["sol_elems"].append(elem)
                else:
                    current_item["q_elems"].append(elem)
            elif current_item is None and in_cluster and current_cluster:
                current_cluster["intro_elems"].append(elem)

    flush_cluster_to_blocks()
    return doc, phan1_blocks, phan2_blocks, phan3_blocks, phan4_blocks

# 7. THIẾT LẬP THÔNG TIN GIAO DIỆN STREAMLIT
with st.expander("⚙️ **THÔNG TIN TIÊU ĐỀ ĐỀ KIỂM TRA & CÁC MÃ ĐỀ**", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        so_gd = st.text_input("Sở Giáo dục & Đào tạo", value="SỞ GIÁO DỤC VÀ ĐÀO TẠO TP.HCM")
        ten_truong = st.text_input("Tên Trường", value="TRƯỜNG THPT THỦ ĐỨC")
        ten_ky_thi = st.text_input("Tên Kỳ Thi", value="KIỂM TRA ĐỊNH KÌ HỌC KÌ I")
    with col2:
        mon_thi = st.text_input("Môn thi / Khối", value="MÔN VẬT LÍ - KHỐI 12")
        thoi_gian = st.text_input("Thời gian làm bài", value="50 phút")
        ds_ma_de_input = st.text_input("Danh sách Mã Đề (cách nhau bởi dấu phẩy)", value="1021, 1022, 1023, 1024")

uploaded_file = st.file_uploader("Bước 1: Tải lên file Word đề gốc (.docx)", type=["docx"])

if uploaded_file is not None:
    st.success(f"✅ Đã nhận file: **{uploaded_file.name}**")

if st.button("🚀 Bước 2: Bắt Đầu Trộn Đề & Tạo File Đáp Án", type="primary"):
    if uploaded_file is None:
        st.warning("⚠️ Vui lòng chọn file Word đề gốc ở Bước 1 trước!")
    else:
        with st.spinner("Đang khởi tạo mã đề thi và xuất file Excel / Word đáp án..."):
            doc_goc_obj, phan1_goc, phan2_goc, phan3_goc, phan4_goc = doc_de_goc(uploaded_file)
            
            danh_sach_ma_de = [m.strip() for m in ds_ma_de_input.split(",") if m.strip()]
            if not danh_sach_ma_de:
                danh_sach_ma_de = ["1021", "1022", "1023", "1024"]
                
            num_p1_total = sum(len(b["items"]) if b.get("is_cluster") else 1 for b in phan1_goc)
            num_p2_total = sum(len(b["items"]) if b.get("is_cluster") else 1 for b in phan2_goc)
            num_p3_total = sum(len(b["items"]) if b.get("is_cluster") else 1 for b in phan3_goc)
            num_p4_total = sum(len(b["items"]) if b.get("is_cluster") else 1 for b in phan4_goc)
            
            tong_so_cau = num_p1_total + num_p2_total + num_p3_total + num_p4_total
            if tong_so_cau == 0:
                st.error("❌ **Không tìm thấy câu hỏi!** Vui lòng kiểm tra lại định dạng đề gốc.")
            else:
                sec_info_parts = []
                if num_p1_total > 0: sec_info_parts.append(f"**{num_p1_total}** câu Phần I")
                if num_p2_total > 0: sec_info_parts.append(f"**{num_p2_total}** câu Phần II")
                if num_p3_total > 0: sec_info_parts.append(f"**{num_p3_total}** câu Phần III")
                if num_p4_total > 0: sec_info_parts.append(f"**{num_p4_total}** câu Phần IV (Tự luận)")
                st.info("📊 Đã phân tích: " + " | ".join(sec_info_parts))
                
                zip_buffer = io.BytesIO()
                ma_de_answers = {}

                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for ma_de in danh_sach_ma_de:
                        doc = Document()
                        Thiet_Lap_Font_Times_New_Roman(doc)
                        rid_map = {}
                        ma_de_answers[ma_de] = {'p1': [], 'p2': [], 'p3': [], 'p4': []}
                        
                        sec_notes = []
                        if num_p1_total > 0: sec_notes.append(f"{num_p1_total} TN")
                        if num_p2_total > 0: sec_notes.append(f"{num_p2_total} ĐS")
                        if num_p3_total > 0: sec_notes.append(f"{num_p3_total} TL Ngắn")
                        if num_p4_total > 0: sec_notes.append(f"{num_p4_total} Tự luận")
                        note_str = f"({', '.join(sec_notes)})" if sec_notes else ""
                        
                        # 1 & 2. TẠO KHUNG HEADER TIÊU ĐỀ HỌ TÊN MÃ ĐỀ
                        Tao_Header_Trang_In(doc, so_gd, ten_truong, ten_ky_thi, mon_thi, thoi_gian, note_str, ma_de)

                        nhan_p1_word = ["A.", "B.", "C.", "D."]
                        nhan_p1_excel = ["A", "B", "C", "D"]
                        nhan_p2_word = ["a)", "b)", "c)", "d)"]

                        # 3. TRỘN PHẦN I
                        if phan1_goc:
                            p_head1 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head1)
                            p_head1.paragraph_format.space_before = Pt(6)
                            p_head1.paragraph_format.space_after = Pt(4)
                            r_h1 = p_head1.add_run(f"PHẦN I. Thí sinh trả lời từ câu 1 đến câu {num_p1_total}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
                            r_h1.bold = True
                            r_h1.font.name = 'Times New Roman'; r_h1.font.size = Pt(12)

                            p1_shuffled_blocks = copy.deepcopy(phan1_goc)
                            random.shuffle(p1_shuffled_blocks)
                            current_q_idx = 1

                            for block in p1_shuffled_blocks:
                                if not block.get("is_cluster"):
                                    q = block["item"]
                                    updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                    append_elem(doc, updated_q0)
                                    
                                    for sub_elem in q["q_elems"][1:]:
                                        sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, sub_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, sub_copied)
                                    
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
                                    current_q_idx += 1
                                else:
                                    c_items = block["items"]
                                    if not block.get("is_fixed", False):
                                        random.shuffle(c_items)
                                    k_count = len(c_items)
                                    start_q = current_q_idx
                                    end_q = current_q_idx + k_count - 1
                                    
                                    Render_Cau_Dan_Chum(doc, start_q, end_q)
                                    
                                    for intro_elem in block["intro_elems"]:
                                        intro_copied = Copy_Element_With_Media(intro_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, intro_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, intro_copied)
                                        
                                    for q in c_items:
                                        updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                        append_elem(doc, updated_q0)
                                        
                                        for sub_elem in q["q_elems"][1:]:
                                            sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                            doc_tmp = Document()
                                            append_elem(doc_tmp, sub_copied)
                                            if len(doc_tmp.paragraphs) > 0:
                                                Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                                append_elem(doc, doc_tmp.paragraphs[0]._element)
                                            else:
                                                append_elem(doc, sub_copied)
                                        
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
                                        current_q_idx += 1

                        # 4. TRỘN PHẦN II
                        if phan2_goc:
                            p_head2 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head2)
                            p_head2.paragraph_format.space_before = Pt(10)
                            p_head2.paragraph_format.space_after = Pt(4)
                            r_h2 = p_head2.add_run(f"PHẦN II. Thí sinh trả lời từ câu 1 đến câu {num_p2_total}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
                            r_h2.bold = True
                            r_h2.font.name = 'Times New Roman'; r_h2.font.size = Pt(12)

                            p2_shuffled_blocks = copy.deepcopy(phan2_goc)
                            random.shuffle(p2_shuffled_blocks)
                            current_q_idx = 1

                            for block in p2_shuffled_blocks:
                                if not block.get("is_cluster"):
                                    q = block["item"]
                                    updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                    append_elem(doc, updated_q0)
                                    
                                    for sub_elem in q["q_elems"][1:]:
                                        sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, sub_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, sub_copied)
                                    
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
                                    current_q_idx += 1
                                else:
                                    c_items = block["items"]
                                    if not block.get("is_fixed", False):
                                        random.shuffle(c_items)
                                    k_count = len(c_items)
                                    start_q = current_q_idx
                                    end_q = current_q_idx + k_count - 1
                                    
                                    Render_Cau_Dan_Chum(doc, start_q, end_q)
                                    
                                    for intro_elem in block["intro_elems"]:
                                        intro_copied = Copy_Element_With_Media(intro_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, intro_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, intro_copied)
                                        
                                    for q in c_items:
                                        updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                        append_elem(doc, updated_q0)
                                        
                                        for sub_elem in q["q_elems"][1:]:
                                            sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                            doc_tmp = Document()
                                            append_elem(doc_tmp, sub_copied)
                                            if len(doc_tmp.paragraphs) > 0:
                                                Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                                append_elem(doc, doc_tmp.paragraphs[0]._element)
                                            else:
                                                append_elem(doc, sub_copied)
                                        
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
                                        current_q_idx += 1

                        # 5. TRỘN PHẦN III
                        if phan3_goc:
                            p_head3 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head3)
                            p_head3.paragraph_format.space_before = Pt(10)
                            p_head3.paragraph_format.space_after = Pt(4)
                            r_h3 = p_head3.add_run(f"PHẦN III. Thí sinh trả lời từ câu 1 đến câu {num_p3_total}.")
                            r_h3.bold = True
                            r_h3.font.name = 'Times New Roman'; r_h3.font.size = Pt(12)

                            p3_shuffled_blocks = copy.deepcopy(phan3_goc)
                            random.shuffle(p3_shuffled_blocks)
                            current_q_idx = 1

                            for block in p3_shuffled_blocks:
                                if not block.get("is_cluster"):
                                    q = block["item"]
                                    updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                    append_elem(doc, updated_q0)
                                    
                                    for sub_elem in q["q_elems"][1:]:
                                        sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, sub_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, sub_copied)
                                        
                                    ma_de_answers[ma_de]['p3'].append(str(q["correct"]).strip())
                                    current_q_idx += 1
                                else:
                                    c_items = block["items"]
                                    if not block.get("is_fixed", False):
                                        random.shuffle(c_items)
                                    k_count = len(c_items)
                                    start_q = current_q_idx
                                    end_q = current_q_idx + k_count - 1
                                    
                                    Render_Cau_Dan_Chum(doc, start_q, end_q)
                                    
                                    for intro_elem in block["intro_elems"]:
                                        intro_copied = Copy_Element_With_Media(intro_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, intro_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, intro_copied)
                                        
                                    for q in c_items:
                                        updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                        append_elem(doc, updated_q0)
                                        
                                        for sub_elem in q["q_elems"][1:]:
                                            sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                            doc_tmp = Document()
                                            append_elem(doc_tmp, sub_copied)
                                            if len(doc_tmp.paragraphs) > 0:
                                                Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                                append_elem(doc, doc_tmp.paragraphs[0]._element)
                                            else:
                                                append_elem(doc, sub_copied)
                                            
                                        ma_de_answers[ma_de]['p3'].append(str(q["correct"]).strip())
                                        current_q_idx += 1

                        # 6. TRỘN PHẦN IV (PHẦN TỰ LUẬN)
                        if phan4_goc:
                            p_head4 = doc.add_paragraph()
                            Set_Paragraph_Justify_And_Zero_Indent(p_head4)
                            p_head4.paragraph_format.space_before = Pt(10)
                            p_head4.paragraph_format.space_after = Pt(4)
                            r_h4 = p_head4.add_run(f"PHẦN IV. Thí sinh trả lời từ câu 1 đến câu {num_p4_total}.")
                            r_h4.bold = True
                            r_h4.font.name = 'Times New Roman'; r_h4.font.size = Pt(12)

                            p4_shuffled_blocks = copy.deepcopy(phan4_goc)
                            random.shuffle(p4_shuffled_blocks)
                            current_q_idx = 1

                            for block in p4_shuffled_blocks:
                                if not block.get("is_cluster"):
                                    q = block["item"]
                                    updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                    append_elem(doc, updated_q0)
                                    
                                    for sub_elem in q["q_elems"][1:]:
                                        sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, sub_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, sub_copied)
                                        
                                    ma_de_answers[ma_de]['p4'].append({"q_num": current_q_idx, "item": q})
                                    current_q_idx += 1
                                else:
                                    c_items = block["items"]
                                    if not block.get("is_fixed", False):
                                        random.shuffle(c_items)
                                    k_count = len(c_items)
                                    start_q = current_q_idx
                                    end_q = current_q_idx + k_count - 1
                                    
                                    Render_Cau_Dan_Chum(doc, start_q, end_q)
                                    
                                    for intro_elem in block["intro_elems"]:
                                        intro_copied = Copy_Element_With_Media(intro_elem, doc_goc_obj, doc, rid_map)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, intro_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc, intro_copied)
                                        
                                    for q in c_items:
                                        updated_q0 = Cap_Nhat_Tien_To_Cau(q["q_elems"][0], doc_goc_obj, doc, rid_map, f"Câu {current_q_idx}: ")
                                        append_elem(doc, updated_q0)
                                        
                                        for sub_elem in q["q_elems"][1:]:
                                            sub_copied = Copy_Element_With_Media(sub_elem, doc_goc_obj, doc, rid_map)
                                            doc_tmp = Document()
                                            append_elem(doc_tmp, sub_copied)
                                            if len(doc_tmp.paragraphs) > 0:
                                                Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                                append_elem(doc, doc_tmp.paragraphs[0]._element)
                                            else:
                                                append_elem(doc, sub_copied)
                                            
                                        ma_de_answers[ma_de]['p4'].append({"q_num": current_q_idx, "item": q})
                                        current_q_idx += 1

                        # 7. LỜI DẶN CHÂN TRANG
                        p_foot1 = doc.add_paragraph()
                        p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foot1.paragraph_format.space_before = Pt(12)
                        p_foot1.paragraph_format.space_after = Pt(2)
                        r_f1 = p_foot1.add_run("---HẾT---")
                        r_f1.bold = True
                        r_f1.font.name = 'Times New Roman'
                        r_f1.font.size = Pt(12)
                        
                        p_foot2 = doc.add_paragraph("Thí sinh không được sử dụng tài liệu - Giám thị không giải thích gì thêm.")
                        p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_f2 = p_foot2.runs[0]
                        r_f2.font.name = 'Times New Roman'; r_f2.font.size = Pt(10); r_f2.italic = True

                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"{ma_de}.docx", doc_io.getvalue())
                    
                    # --- XUẤT FILE EXCEL ĐÁP ÁN MATRẬN PHẦN I, II, III ---
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

                    # --- XUẤT FILE WORD ĐÁP ÁN PHẦN IV TỰ LUẬN (CHỈ KHI ĐỀ GỐC CÓ PHẦN IV) ---
                    if num_p4_total > 0:
                        doc_p4_ans = Document()
                        Thiet_Lap_Font_Times_New_Roman(doc_p4_ans)
                        rid_map_ans = {}

                        for md_idx, ma_de in enumerate(danh_sach_ma_de):
                            if md_idx > 0:
                                doc_p4_ans.add_page_break()

                            Tao_Header_Trang_In(doc_p4_ans, so_gd, ten_truong, ten_ky_thi, mon_thi, thoi_gian, "", ma_de)

                            p_title_ans = doc_p4_ans.add_paragraph()
                            p_title_ans.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_title_ans.paragraph_format.space_before = Pt(8)
                            p_title_ans.paragraph_format.space_after = Pt(8)
                            r_t_ans = p_title_ans.add_run(f"ĐÁP ÁN & LỜI GIẢI CHI TIẾT PHẦN TỰ LUẬN - MÃ ĐỀ: {ma_de}")
                            r_t_ans.bold = True
                            r_t_ans.font.name = 'Times New Roman'
                            r_t_ans.font.size = Pt(13)

                            for q_info in ma_de_answers[ma_de]['p4']:
                                q_num = q_info["q_num"]
                                q_item = q_info["item"]
                                sol_elems = q_item.get("sol_elems", [])

                                p_qnum = doc_p4_ans.add_paragraph()
                                Set_Paragraph_Justify_And_Zero_Indent(p_qnum)
                                p_qnum.paragraph_format.space_before = Pt(6)
                                p_qnum.paragraph_format.space_after = Pt(2)
                                r_q = p_qnum.add_run(f"Câu {q_num}:")
                                r_q.bold = True
                                r_q.font.name = 'Times New Roman'
                                r_q.font.size = Pt(12)

                                if sol_elems:
                                    for sol_elem in sol_elems:
                                        sol_copied = Copy_Element_With_Media(sol_elem, doc_goc_obj, doc_p4_ans, rid_map_ans)
                                        doc_tmp = Document()
                                        append_elem(doc_tmp, sol_copied)
                                        if len(doc_tmp.paragraphs) > 0:
                                            Set_Paragraph_Justify_And_Zero_Indent(doc_tmp.paragraphs[0])
                                            append_elem(doc_p4_ans, doc_tmp.paragraphs[0]._element)
                                        else:
                                            append_elem(doc_p4_ans, sol_copied)
                                else:
                                    p_empty = doc_p4_ans.add_paragraph()
                                    Set_Paragraph_Justify_And_Zero_Indent(p_empty)
                                    p_empty.paragraph_format.space_before = Pt(0)
                                    p_empty.paragraph_format.space_after = Pt(4)
                                    r_e = p_empty.add_run("(Chưa có lời giải chi tiết)")
                                    r_e.italic = True
                                    r_e.font.name = 'Times New Roman'
                                    r_e.font.size = Pt(12)

                        doc_p4_io = io.BytesIO()
                        doc_p4_ans.save(doc_p4_io)
                        zip_file.writestr("Dap_an_Tu_luan_cac_Ma_de.docx", doc_p4_io.getvalue())

                st.success("🎉 Trộn đề thành công!")
                st.download_button(
                    label="📥 TẢI XUỐNG CÁC MÃ ĐỀ & FILE ĐÁP ÁN (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="DE_DA_TRON.zip",
                    mime="application/zip",
                    type="primary"
                )

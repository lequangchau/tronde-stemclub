import streamlit as st
import os
import re
import random
import zipfile
import io
import copy
from docx import Document
import pandas as pd

# 1. CẤU HÌNH TRANG WEB STREAMLIT
st.set_page_config(page_title="Phần mềm Trộn Đề Vật Lí", page_icon="📝", layout="centered")

# 2. GIAO DIỆN CHÍNH & THÔNG TIN TÁC GIẢ
st.title("📝 Ứng Dụng Trộn Đề Vật Lí Online")
st.caption("👨‍🏫 **Tác giả:** LQChau - CLB STEM - THPT Thủ Đức")
st.write("Hệ thống hỗ trợ trộn đề 3 phần (THPT Quốc Gia) và xuất đáp án sang Excel.")

st.divider()

# 3. HÀM ĐỌC VÀ PHÂN TÍCH ĐỀ GỐC FILE WORD (.DOCX)
def doc_de_goc(file_stream):
    doc = Document(file_stream)
    phan1, phan2, phan3 = [], [], []
    current_section = None
    current_item = None

    for p in doc.paragraphs:
        line = p.text.strip()
        if not line:
            continue

        # Nhận diện phân đoạn
        if "[PHAN_1]" in line:
            current_section = 1
            continue
        elif "[PHAN_2]" in line:
            if current_item:
                phan1.append(current_item)
                current_item = None
            current_section = 2
            continue
        elif "[PHAN_3]" in line:
            if current_item:
                phan2.append(current_item)
                current_item = None
            current_section = 3
            continue

        # PHẦN I: Trắc nghiệm 4 lựa chọn
        if current_section == 1:
            if re.match(r'^(Câu|Cau)\s*\d+[\.\:]', line, re.IGNORECASE):
                if current_item:
                    phan1.append(current_item)
                current_item = {"question": line, "options": [], "correct": None}
            elif re.match(r'^[A-D][\.\:]', line, re.IGNORECASE) and current_item:
                clean_opt = re.sub(r'^[A-D][\.\:]\s*', '', line)
                current_item["options"].append(clean_opt)
                is_underlined = any(bool(r.underline) for r in p.runs if r.text.strip())
                if is_underlined:
                    current_item["correct"] = clean_opt

        # PHẦN II: Trắc nghiệm Đúng/Sai
        elif current_section == 2:
            if re.match(r'^(Câu|Cau)\s*\d+[\.\:]', line, re.IGNORECASE):
                if current_item:
                    phan2.append(current_item)
                current_item = {"question": line, "sub_options": []}
            elif re.match(r'^[a-d][\)\.\:]', line, re.IGNORECASE) and current_item:
                clean_sub = re.sub(r'^[a-d][\)\.\:]\s*', '', line)
                is_underlined = any(bool(r.underline) for r in p.runs if r.text.strip())
                current_item["sub_options"].append({
                    "text": clean_sub, 
                    "is_correct": "Đúng" if is_underlined else "Sai"
                })

        # PHẦN III: Trả lời ngắn
        elif current_section == 3:
            if re.match(r'^(Câu|Cau)\s*\d+[\.\:]', line, re.IGNORECASE):
                if current_item:
                    phan3.append(current_item)
                current_item = {"question": line, "correct": ""}
            elif "Đáp số:" in line and current_item:
                current_item["correct"] = line.replace("Đáp số:", "").strip()
    
    if current_item:
        if current_section == 1:
            phan1.append(current_item)
        elif current_section == 2:
            phan2.append(current_item)
        elif current_section == 3:
            phan3.append(current_item)

    return phan1, phan2, phan3

# 4. XỬ LÝ TRỘN ĐỀ VÀ TẠO BỘ ĐỀ
uploaded_file = st.file_uploader("Bước 1: Tải lên file Word đề gốc (.docx)", type=["docx"])
so_ma_de = st.number_input("Bước 2: Nhập số lượng mã đề muốn tạo", min_value=1, max_value=20, value=4, step=1)

if uploaded_file is not None:
    if st.button("Bước 3: Bắt Đầu Trộn Đề & Sinh Đáp Án"):
        with st.spinner("Đang xử lý dữ liệu..."):
            phan1_goc, phan2_goc, phan3_goc = doc_de_goc(uploaded_file)
            
            zip_buffer = io.BytesIO()
            nhan_p1 = ["A.", "B.", "C.", "D."]
            nhan_p2 = ["a)", "b)", "c)", "d)"]
            danh_sach_excel_data = []

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                for i in range(so_ma_de):
                    ma_de = 201 + i
                    doc = Document()
                    doc.add_heading(f"ĐỀ THI MINH HỌA VẬT LÍ - MÃ ĐỀ: {ma_de}", level=1)
                    
                    # --- TRỘN PHẦN I ---
                    doc.add_heading("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.", level=2)
                    p1_shuffled = copy.deepcopy(phan1_goc)
                    random.shuffle(p1_shuffled)
                    for idx, q in enumerate(p1_shuffled):
                        correct_text = q["correct"]
                        random.shuffle(q["options"])
                        txt_cau = re.sub(r'^(Câu|Cau)\s*\d+[\.\:]\s*', '', q["question"], flags=re.IGNORECASE)
                        doc.add_paragraph(f"Câu {idx + 1}. {txt_cau}")
                        
                        correct_label = "Chưa chọn đáp án đúng ở đề gốc"
                        for o_idx, opt in enumerate(q["options"]):
                            doc.add_paragraph(f"   {nhan_p1[o_idx]} {opt}")
                            if correct_text is not None and opt == correct_text:
                                correct_label = nhan_p1[o_idx]
                        
                        danh_sach_excel_data.append({
                            "Mã Đề": ma_de, 
                            "Phần": "Phần I", 
                            "Câu Số": idx + 1, 
                            "Đáp Án Đúng": correct_label
                        })

                    # --- TRỘN PHẦN II ---
                    doc.add_heading("PHẦN II. Câu trắc nghiệm đúng sai.", level=2)
                    p2_shuffled = copy.deepcopy(phan2_goc)
                    random.shuffle(p2_shuffled)
                    for idx, q in enumerate(p2_shuffled):
                        random.shuffle(q["sub_options"])
                        txt_cau = re.sub(r'^(Câu|Cau)\s*\d+[\.\:]\s*', '', q["question"], flags=re.IGNORECASE)
                        doc.add_paragraph(f"Câu {idx + 1}. {txt_cau}")
                        
                        for s_idx, sub in enumerate(q["sub_options"]):
                            doc.add_paragraph(f"   {nhan_p2[s_idx]} {sub['text']}")
                            danh_sach_excel_data.append({
                                "Mã Đề": ma_de, 
                                "Phần": "Phần II", 
                                "Câu Số": f"Câu {idx + 1} ý {nhan_p2[s_idx]}", 
                                "Đáp Án Đúng": sub['is_correct']
                            })

                    # --- TRỘN PHẦN III ---
                    doc.add_heading("PHẦN III. Câu trắc nghiệm trả lời ngắn.", level=2)
                    p3_shuffled = copy.deepcopy(phan3_goc)
                    random.shuffle(p3_shuffled)
                    for idx, q in enumerate(p3_shuffled):
                        txt_cau = re.sub(r'^(Câu|Cau)\s*\d+[\.\:]\s*', '', q["question"], flags=re.IGNORECASE)
                        doc.add_paragraph(f"Câu {idx + 1}. {txt_cau}")
                        doc.add_paragraph("   Đáp số: ................................")
                        danh_sach_excel_data.append({
                            "Mã Đề": ma_de, 
                            "Phần": "Phần III", 
                            "Câu Số": idx + 1, 
                            "Đáp Án Đúng": q["correct"]
                        })

                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    zip_file.writestr(f"De_Thi_Ma_De_{ma_de}.docx", doc_io.getvalue())
                
                # --- TẠO FILE EXCEL ĐÁP ÁN TỔNG HỢP ---
                df = pd.DataFrame(danh_sach_excel_data)
                excel_io = io.BytesIO()
                with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="Dap_An_Tong_Hop")
                zip_file.writestr("Bang_Dap_An_Cac_Ma_De.xlsx", excel_io.getvalue())

            st.success("🎉 Trộn đề hoàn tất thành công!")
            st.download_button(
                label="📥 TẢI XUỐNG TOÀN BỘ ĐỀ THI & FILE EXCEL ĐÁP ÁN (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Bo_De_Thi_Va_Dap_An.zip",
                mime="application/zip"
            )

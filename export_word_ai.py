from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def tao_file_word_de_thi(config, output_filename, generated_questions=None, kem_dap_an=True):
    doc=Document()
    normal=doc.styles["Normal"]; normal.font.name="Times New Roman"; normal.font.size=Pt(12)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f"ĐỀ KIỂM TRA {config['loai_kiem_tra'].upper()}\\n"); r.bold=True; r.font.size=Pt(14)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Môn: {config['mon_hoc']} - Khối: {config['khoi_lop']}\\nThời gian: {config['thoi_gian']} phút")
    doc.add_paragraph("Họ và tên: .................................................... Lớp: ...............")
    questions=generated_questions or []
    groups=[("Dang_I","PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN"),("Dang_II","PHẦN II. TRẮC NGHIỆM ĐÚNG/SAI"),("Dang_III","PHẦN III. TỰ LUẬN")]
    n=1
    for kind,title in groups:
        qs=[q for q in questions if q["dang_cau_hoi"]==kind]
        if not qs: continue
        doc.add_paragraph(title).runs[0].bold=True
        for q in qs:
            p=doc.add_paragraph(); p.add_run(f"Câu {n}. ").bold=True; p.add_run(q.get("noi_dung") or "")
            if kind=="Dang_I":
                doc.add_paragraph(f"A. {q.get('dap_an_a') or ''}     B. {q.get('dap_an_b') or ''}")
                doc.add_paragraph(f"C. {q.get('dap_an_c') or ''}     D. {q.get('dap_an_d') or ''}")
            elif kind=="Dang_II":
                for label,key in [("a","dap_an_a"),("b","dap_an_b"),("c","dap_an_c"),("d","dap_an_d")]:
                    doc.add_paragraph(f"{label}) {q.get(key) or ''}")
            n+=1
    if kem_dap_an:
        doc.add_page_break(); doc.add_paragraph("ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM").runs[0].bold=True
        n=1
        for kind,_ in groups:
            for q in [x for x in questions if x["dang_cau_hoi"]==kind]:
                ans=(q.get("dap_an_a") or q.get("dap_an_dung") or "") if kind=="Dang_III" else (q.get("dap_an_dung") or "")
                p=doc.add_paragraph(); p.add_run(f"Câu {n}: ").bold=True; p.add_run(ans)
                if q.get("giai_thich"): p.add_run("\\nGiải thích/Hướng dẫn chấm: "+q["giai_thich"])
                n+=1
    doc.save(output_filename)
    return output_filename

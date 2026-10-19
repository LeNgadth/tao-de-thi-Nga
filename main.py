import json
import os
import uuid
from pathlib import Path
from io import BytesIO

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from openai import OpenAI
from pypdf import PdfReader
from export_word_ai import tao_file_word_de_thi

app = FastAPI(title="Hệ thống tạo đề thi bằng AI")
from fastapi.responses import FileResponse
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

@app.get("/")
def home():
    return FileResponse(BASE_DIR / "index.html")

def read_ppct(file_ppct):
    if not file_ppct or not file_ppct.filename:
        return ""
    if not file_ppct.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Vui lòng tải file PPCT định dạng PDF.")
    try:
        reader = PdfReader(BytesIO(file_ppct.file.read()))
        return "\\n".join(page.extract_text() or "" for page in reader.pages).strip()[:30000]
    except Exception as exc:
        raise HTTPException(400, f"Không đọc được PDF: {exc}")


def generate_questions(mon_hoc, khoi_lop, loai_kiem_tra, chu_de, thoi_gian, ma_tran, ppct):
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(500, "Chưa cấu hình OPENAI_API_KEY trong Environment Variables của Render.")
    requested = [
        {"dang_cau_hoi": dang, "muc_do": muc, "so_luong": n}
        for dang, mucs in ma_tran.items() for muc, n in mucs.items() if n > 0
    ]
    if not requested:
        raise HTTPException(400, "Vui lòng nhập số lượng ít nhất 1 câu hỏi.")
    prompt = f'''
Tạo đề kiểm tra THPT bằng tiếng Việt.
Môn: {mon_hoc}; khối: {khoi_lop}; loại: {loai_kiem_tra}; thời gian: {thoi_gian} phút.
Chủ đề/phạm vi giáo viên nhập: {chu_de or "Bám sát PPCT hoặc kiến thức phù hợp"}.
Nội dung PPCT trích từ PDF: {ppct or "Không có PDF"}.
Ma trận số câu bắt buộc: {json.dumps(requested, ensure_ascii=False)}.
Chỉ trả JSON dạng {{"questions":[...]}}. Mỗi câu có các trường:
dang_cau_hoi (Dang_I hoặc Dang_II), muc_do (biet/hieu/vandung), noi_dung,
dap_an_a, dap_an_b, dap_an_c, dap_an_d, dap_an_dung, giai_thich.
Dang_I: 4 lựa chọn, dap_an_dung là A/B/C/D.
Dang_II: nội dung là hướng dẫn chung; dap_an_a-d là 4 mệnh đề; dap_an_dung ghi 4 đáp án theo thứ tự, ví dụ "Đúng,Sai,Đúng,Sai".
Dang_III: câu tự luận; noi_dung là đề bài; dap_an_a ghi đáp án mẫu/hướng dẫn chấm; dap_an_dung ghi đáp án tóm tắt.
Tạo đúng tổng số câu theo ma trận, không thêm/bớt; câu rõ ràng, không trùng lặp, đáp án chính xác, phù hợp khối lớp. Không dùng markdown.
'''
    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            temperature=0.3,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "Bạn là chuyên gia biên soạn đề kiểm tra. Chỉ trả JSON hợp lệ."},
                {"role": "user", "content": prompt},
            ],
        )
        questions = json.loads(response.choices[0].message.content or "{}").get("questions", [])
    except Exception as exc:
        raise HTTPException(502, f"AI không tạo được đề: {exc}")

    for q in questions:
        if not all(q.get(k) for k in ("dang_cau_hoi", "muc_do", "noi_dung")):
            raise HTTPException(502, "AI trả về câu hỏi thiếu dữ liệu. Vui lòng thử lại.")
        for k in ("dap_an_a", "dap_an_b", "dap_an_c", "dap_an_d", "dap_an_dung", "giai_thich"):
            q.setdefault(k, "")
    expected = sum(n for group in ma_tran.values() for n in group.values())
    if len(questions) != expected:
        raise HTTPException(502, f"AI trả về {len(questions)} câu, ma trận yêu cầu {expected}. Vui lòng tạo lại.")
    return questions


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/tao-ma-tran")
async def tao_ma_tran_de_thi(
    mon_hoc: str = Form(...),
    khoi_lop: str = Form(...),
    loai_kiem_tra: str = Form(...),
    thoi_gian: int = Form(...),
    chu_de: str = Form(""),
    dang1_biet: int = Form(0),
    dang1_hieu: int = Form(0),
    dang1_vandung: int = Form(0),
    dang2_biet: int = Form(0),
    dang2_hieu: int = Form(0),
    dang2_vandung: int = Form(0),
    dang3_biet: int = Form(0),
    dang3_hieu: int = Form(0),
    dang3_vandung: int = Form(0),
    file_ppct: UploadFile | None = File(None),
):
    if thoi_gian <= 0:
        raise HTTPException(400, "Thời gian làm bài phải lớn hơn 0.")
    counts = [
        dang1_biet, dang1_hieu, dang1_vandung,
        dang2_biet, dang2_hieu, dang2_vandung,
        dang3_biet, dang3_hieu, dang3_vandung,
    ]
    if any(n < 0 for n in counts):
        raise HTTPException(400, "Số lượng câu hỏi không được âm.")
    ma_tran = {
        "Dang_I": {"biet": dang1_biet, "hieu": dang1_hieu, "vandung": dang1_vandung},
        "Dang_II": {"biet": dang2_biet, "hieu": dang2_hieu, "vandung": dang2_vandung},
        "Dang_III": {"biet": dang3_biet, "hieu": dang3_hieu, "vandung": dang3_vandung},
    }
    ppct = read_ppct(file_ppct)
    questions = generate_questions(mon_hoc, khoi_lop, loai_kiem_tra, chu_de, thoi_gian, ma_tran, ppct)
    config = {
        "mon_hoc": mon_hoc, "khoi_lop": khoi_lop, "loai_kiem_tra": loai_kiem_tra,
        "thoi_gian": thoi_gian,
        "ma_tran": {
            "dang_1_nhieu_lua_chon": ma_tran["Dang_I"],
            "dang_2_dung_sai": ma_tran["Dang_II"],
            "dang_3_tu_luan": ma_tran["Dang_III"],
        },
    }
    output = str(Path("/tmp") / f"De_Thi_{uuid.uuid4().hex[:8]}.docx")
    try:
        tao_file_word_de_thi(config, output, generated_questions=questions)
    except Exception as exc:
        raise HTTPException(500, f"Không xuất được Word: {exc}")
    return FileResponse(
        output, filename=f"De_Thi_Khoi_{khoi_lop}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

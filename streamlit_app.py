# streamlit_app.py
import streamlit as st
import pandas as pd
import io
import re
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from reportlab.lib import colors
from functools import lru_cache

# -----------------------
# 폰트 (자동 대체 + NotoSansKR)
# -----------------------
FONT_NAME = None

def register_korean_font():
    global FONT_NAME
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
        FONT_NAME = "HYSMyeongJo-Medium"
    except:
        try:
            os.makedirs("fonts", exist_ok=True)
            path = "fonts/NotoSansKR-Regular.ttf"
            if not os.path.exists(path):
                st.info("NotoSansKR 다운로드 중...")
                import urllib.request
                urllib.request.urlretrieve(
                    "https://github.com/google/fonts/raw/main/ofl/notosanskr/NotoSansKR-Regular.ttf",
                    path
                )
            pdfmetrics.registerFont(TTFont("NotoSansKR", path))
            FONT_NAME = "NotoSansKR"
        except:
            FONT_NAME = "Helvetica"
            st.warning("한글 폰트 없음 → Helvetica 사용 (한글 깨짐 가능)")

register_korean_font()

# -----------------------
# 상수 (mm)
# -----------------------
NUM_X1_MM = 24; UNDER_X1_MM = 62; UNDER_LEN_MM = 37
NUM_X2_MM = 111; UNDER_X2_MM = 152
TOP_OFFSET_MM = 62; BOTTOM_RESERVED_MM = 24
LINE_HEIGHT_MM = 4; GAP_AFTER_MM = 4
CHAR_SIZE_MM = 2

# -----------------------
# 캐시 + 줄바꿈
# -----------------------
@lru_cache(maxsize=10000)
def cached_string_width(text, font_name, font_size_pt):
    return pdfmetrics.stringWidth(text, font_name, font_size_pt)

def wrap_text(text, font_size_pt, max_width_pt):
    if not text or not text.strip():
        return [""]
    lines = []
    cur = ""
    for ch in text:
        cand = cur + ch
        if cached_string_width(cand, FONT_NAME, font_size_pt) <= max_width_pt:
            cur = cand
        else:
            if cur: lines.append(cur)
            cur = ch
    if cur: lines.append(cur)
    return lines

# -----------------------
# Day 추출
# -----------------------
def extract_day_label(name):
    name = name.rsplit("/", 1)[-1]
    name = ".".join(name.split(".")[:-1]) if "." in name else name
    m = re.search(r"Day[-_ ]*(\d+)", name, re.I)
    return f"Day {int(m.group(1))}" if m else name

# -----------------------
# 페이지 수 계산 (정확)
# -----------------------
def simulate_page_count(word_pairs, n):
    if n == 0: return 1
    page_h = A4[1]
    top_y = page_h - TOP_OFFSET_MM * mm
    bottom_limit = BOTTOM_RESERVED_MM * mm
    line_h = LINE_HEIGHT_MM * mm
    gap = GAP_AFTER_MM * mm
    needed = line_h + gap

    y_left = top_y
    y_right = top_y
    pages = 1

    for eng, kor, is_eng in word_pairs[:n]:
        text = eng if is_eng else kor
        max_w = (UNDER_X1_MM * mm - 2 * mm) - (NUM_X1_MM * mm + 6 * mm)
        lines = wrap_text(text, CHAR_SIZE_MM * mm, max_w)
        block_h = len(lines) * line_h
        total_h = block_h + gap

        if is_eng:
            if y_left - total_h < bottom_limit:
                pages += 1
                y_left = top_y
                y_right = top_y
            y_left -= total_h
        else:
            if y_right - total_h < bottom_limit:
                pages += 1
                y_left = top_y
                y_right = top_y
            y_right -= total_h
    return pages

# -----------------------
# PDF 생성 (시험지)
# -----------------------
def create_pdf(word_pairs, n, label, total_pages, is_answer=False):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # 좌표
    num_x1 = NUM_X1_MM * mm; under_x1 = UNDER_X1_MM * mm; under_len = UNDER_LEN_MM * mm
    num_x2 = NUM_X2_MM * mm; under_x2 = UNDER_X2_MM * mm
    top_y = h - TOP_OFFSET_MM * mm
    bottom_limit = BOTTOM_RESERVED_MM * mm
    line_h = LINE_HEIGHT_MM * mm
    gap = GAP_AFTER_MM * mm
    font_size = CHAR_SIZE_MM * mm

    text_x1 = num_x1 + 6 * mm; text_x2 = num_x2 + 6 * mm
    max_w1 = (under_x1 - 2 * mm) - text_x1
    max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, font_size)

    idx = 0
    page_no = 1

    while idx < n:
        if page_no == 1:
            # 헤더
            header_y = h - 6 * mm
            c.setFont(FONT_NAME, font_size * 2.5)
            c.drawCentredString(w / 2, header_y, f"{label} - 영어 단어 시험지")
            c.setLineWidth(0.5)
            c.line(12 * mm, header_y - 4 * mm, w - 12 * mm, header_y - 4 * mm)
            meta_y = header_y - 10 * mm
            c.setFont(FONT_NAME, font_size)
            x = w / 2 + 10 * mm
            c.drawString(x, meta_y, "이름: _______________________")
            c.drawString(x, meta_y - 6 * mm, "학년/반: __ / __     번호: ____")
            c.drawString(x, meta_y - 12 * mm, "점수: ______ / 100")
            c.drawString(x, meta_y - 18 * mm, "시험일: ____________________")
            c.setFont(FONT_NAME, font_size)

        y_left = top_y
        y_right = top_y

        while idx < n:
            eng, kor, is_eng = word_pairs[idx]
            text = eng if is_eng else kor
            max_w = max_w1 if is_eng else max_w2
            lines = wrap_text(text, font_size, max_w)
            block_h = len(lines) * line_h
            total_h = block_h + gap

            if is_eng:
                if y_left - total_h < bottom_limit: break
                y_cur = y_left
                num_x = num_x1; txt_x = text_x1; und_x = under_x1
                y_left -= total_h
            else:
                if y_right - total_h < bottom_limit: break
                y_cur = y_right
                num_x = num_x2; txt_x = text_x2; und_x = under_x2
                y_right -= total_h

            # 출력
            c.setFillColor(colors.black)
            c.drawString(num_x, y_cur, f"{idx+1}.")
            for i, line in enumerate(lines):
                c.drawString(txt_x, y_cur - i * line_h, line)
            c.setLineWidth(0.5)
            c.line(und_x, y_cur - 0.1 * mm, und_x + under_len, y_cur - 0.1 * mm)

            # 정답지
            if is_answer:
                c.setFillColor(colors.blue)
                ans = kor if is_eng else eng
                ans_lines = wrap_text(ans, font_size, under_len - 2 * mm)
                for i, a in enumerate(ans_lines):
                    c.drawString(und_x + 1 * mm, y_cur - i * line_h, a)
                c.setFillColor(colors.black)

            idx += 1

        # 페이지 번호
        c.setFont(FONT_NAME, font_size)
        c.drawCentredString(w / 2, bottom_limit / 2, f"Page {page_no} / {total_pages}")

        if idx < n:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(layout="wide", page_title="영어 시험지")
st.title("영어 단어 시험지 생성기")
st.write("영어 왼쪽 열, 한글 오른쪽 열, 페이지당 최대 배치")

uploaded_files = st.file_uploader("파일 업로드", type=["xlsx", "csv"], accept_multiple_files=True)
num_questions = st.number_input("문항 수", min_value=2, max_value=500, value=60, step=2)

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        try:
            df = pd.read_excel(f) if f.name.endswith(".xlsx") else pd.read_csv(io.BytesIO(f.read()))
            df.columns = [c.strip().lower() for c in df.columns]
            if "english" in df.columns and "korean" in df.columns:
                df = df[["english", "korean"]]
            elif "단어" in df.columns and "뜻" in df.columns:
                df = df[["단어", "뜻"]].rename(columns={"단어": "english", "뜻": "korean"})
            else:
                df = df.iloc[:, :2].rename(columns={df.columns[0]: "english", df.columns[1]: "korean"})
            dfs.append(df)
        except Exception as e:
            st.error(f"{f.name}: {e}")

    if not dfs:
        st.warning("데이터 없음")
    else:
        df = pd.concat(dfs).dropna().drop_duplicates("english")
        n = min(num_questions, len(df))
        sampled = df.sample(frac=1, random_state=42).iloc[:n]

        # Day 라벨
        labels = [extract_day_label(f.name) for f in uploaded_files]
        labels = sorted(set(labels), key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999)
        label = " - ".join(labels) if labels else "영어 단어 시험지"

        # word_pairs
        half = n // 2
        pairs = [(str(sampled.iloc[i]["english"]), str(sampled.iloc[i]["korean"]), True) for i in range(half)]
        pairs += [(str(sampled.iloc[i]["english"]), str(sampled.iloc[i]["korean"]), False) for i in range(half, n)]

        total_pages = simulate_page_count(pairs, n)

        test_buf = create_pdf(pairs, n, label, total_pages, is_answer=False)
        ans_buf = create_pdf(pairs, n, label, total_pages, is_answer=True)

        st.download_button("시험지 다운로드", test_buf, "시험지.pdf", "application/pdf")
        st.download_button("정답지 다운로드", ans_buf, "정답지.pdf", "application/pdf")
        st.success(f"완료! {n}문항, {total_pages}페이지")
else:
    st.info("파일 업로드하세요.")

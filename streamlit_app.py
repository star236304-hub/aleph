# streamlit_app.py
import streamlit as st
import pandas as pd
import io
import re
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from reportlab.lib import colors
import os
from functools import lru_cache

# -----------------------
# 폰트 안전 로딩 (자동 대체 + NotoSansKR 다운로드)
# -----------------------
FONT_NAME = None

def register_korean_font():
    global FONT_NAME
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
        FONT_NAME = "HYSMyeongJo-Medium"
        st.success("폰트: HYSMyeongJo-Medium 사용")
    except:
        try:
            os.makedirs("fonts", exist_ok=True)
            font_path = "fonts/NotoSansKR-Regular.ttf"
            if not os.path.exists(font_path):
                st.info("NotoSansKR 폰트 다운로드 중...")
                import urllib.request
                url = "https://github.com/google/fonts/raw/main/ofl/notosanskr/NotoSansKR-Regular.ttf"
                urllib.request.urlretrieve(url, font_path)
            pdfmetrics.registerFont(TTFont("NotoSansKR", font_path))
            FONT_NAME = "NotoSansKR"
            st.success("폰트: NotoSansKR 사용 (한글+영어 완벽)")
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
LINE_HEIGHT_EN_MM = 4; LINE_HEIGHT_KO_GAP_MM = 4
CHAR_SIZE_MM = 2

# -----------------------
# 캐싱 + 줄바꿈
# -----------------------
@lru_cache(maxsize=10000)
def cached_string_width(text, font_name, font_size_pt):
    return pdfmetrics.stringWidth(text, font_name, font_size_pt)

def wrap_text_by_width(text, font_name, font_size_pt, max_width_pt):
    if not text or not text.strip():
        return [""]
    lines = []
    cur = ""
    for ch in text:
        candidate = cur + ch
        if cached_string_width(candidate, font_name, font_size_pt) <= max_width_pt:
            cur = candidate
        else:
            if cur:
                lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines

# -----------------------
# Day 추출
# -----------------------
def extract_day_label(filename):
    name = filename.rsplit("/", 1)[-1]
    name_wo_ext = ".".join(name.split(".")[:-1]) if "." in name else name
    m = re.search(r"Day[-_ ]*(\d+)", name_wo_ext, flags=re.IGNORECASE)
    return f"Day {int(m.group(1))}" if m else name_wo_ext

# -----------------------
# 페이지 수 계산 (정확)
# -----------------------
def simulate_page_count(word_pairs, num_questions):
    total = min(num_questions, len(word_pairs))
    if total == 0:
        return 1

    page_w_pt, page_h_pt = A4
    top_y = page_h_pt - TOP_OFFSET_MM * mm
    bottom_y = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    per_line_h = char_size_pt * 1.1
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    y_left = top_y
    y_right = top_y
    page_count = 1

    for i in range(total):
        eng, kor, is_kor_blank = word_pairs[i]
        shown_text = eng if is_kor_blank else kor
        is_english = is_kor_blank

        # 열 선택
        if is_english and y_left > y_right:
            y_current = y_left
            text_max_w = (UNDER_X1_MM * mm - 2 * mm) - (NUM_X1_MM * mm + 6 * mm)
        else:
            y_current = y_right
            text_max_w = (UNDER_X2_MM * mm - 2 * mm) - (NUM_X2_MM * mm + 6 * mm)

        lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
        block_h = len(lines) * per_line_h
        gap_after = line_height_en if is_kor_blank else ko_gap
        needed = block_h + gap_after

        if y_current - needed < bottom_y:
            page_count += 1
            y_left = top_y
            y_right = top_y
            continue

        if is_english and y_left > y_right:
            y_left = y_current - needed
        else:
            y_right = y_current - needed

    return page_count

# -----------------------
# PDF 생성 (시험지)
# -----------------------
def create_test_pdf(word_pairs, num_questions, filename_label=None, total_pages=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    num_x1 = NUM_X1_MM * mm; under_x1 = UNDER_X1_MM * mm; under_len = UNDER_LEN_MM * mm
    num_x2 = NUM_X2_MM * mm; under_x2 = UNDER_X2_MM * mm
    top_y = page_h_pt - TOP_OFFSET_MM * mm
    bottom_y = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    per_line_h = char_size_pt * 1.1
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    text_x1 = num_x1 + 6 * mm; text_x2 = num_x2 + 6 * mm
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, char_size_pt)

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1

    while idx < total:
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_left = top_y
        y_right = top_y

        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            is_english = is_kor_blank

            # 열 선택: 영어는 왼쪽 우선, 한글은 오른쪽 우선
            if is_english and y_left >= y_right:
                y_current = y_left
                num_x = num_x1; text_x = text_x1; under_x = under_x1; text_max_w = text_max_w1
            elif not is_english and y_right >= y_left:
                y_current = y_right
                num_x = num_x2; text_x = text_x2; under_x = under_x2; text_max_w = text_max_w2
            else:
                break  # 다음 페이지

            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
            if not lines:
                lines = [""]
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed = block_h + gap_after

            if y_current - needed < bottom_y:
                break

            # 배치
            c.setFillColor(colors.black)
            c.drawString(num_x, y_current, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_current - li * per_line_h
                c.drawString(text_x, line_y, line)
            underline_y = y_current - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.5)
            c.line(under_x, underline_y, under_x + under_len, underline_y)

            # y 위치 갱신
            if is_english and y_left >= y_right:
                y_left = y_current - needed
            else:
                y_right = y_current - needed

            idx += 1

        # 페이지 번호
        page_text = f"Page {page_no}"
        if total_pages:
            page_text += f" / {total_pages}"
        c.setFont(FONT_NAME, char_size_pt)
        c.drawCentredString(page_w_pt / 2, bottom_y / 2, page_text)

        if idx < total:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# 정답지 생성 (동일 로직)
# -----------------------
def create_answer_pdf(word_pairs, num_questions, filename_label=None, total_pages=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    num_x1 = NUM_X1_MM * mm; under_x1 = UNDER_X1_MM * mm; under_len = UNDER_LEN_MM * mm
    num_x2 = NUM_X2_MM * mm; under_x2 = UNDER_X2_MM * mm
    top_y = page_h_pt - TOP_OFFSET_MM * mm
    bottom_y = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    per_line_h = char_size_pt * 1.1
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    text_x1 = num_x1 + 6 * mm; text_x2 = num_x2 + 6 * mm
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, char_size_pt)

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1

    while idx < total:
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_left = top_y
        y_right = top_y

        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            is_english = is_kor_blank

            if is_english and y_left >= y_right:
                y_current = y_left
                num_x = num_x1; text_x = text_x1; under_x = under_x1; text_max_w = text_max_w1
            elif not is_english and y_right >= y_left:
                y_current = y_right
                num_x = num_x2; text_x = text_x2; under_x = under_x2; text_max_w = text_max_w2
            else:
                break

            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
            if not lines:
                lines = [""]
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed = block_h + gap_after

            if y_current - needed < bottom_y:
                break

            c.setFillColor(colors.black)
            c.drawString(num_x, y_current, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_current - li * per_line_h
                c.drawString(text_x, line_y, line)
            underline_y = y_current - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.line(under_x, underline_y, under_x + under_len, underline_y)

            # 정답 (파란색)
            c.setFillColor(colors.blue)
            answer_text = kor if is_kor_blank else eng
            answer_lines = wrap_text_by_width(answer_text, FONT_NAME, char_size_pt, under_len - 2*mm)
            for li, a_line in enumerate(answer_lines):
                a_y = y_current - li * per_line_h
                c.drawString(under_x + 1*mm, a_y, a_line)
            c.setFillColor(colors.black)

            if is_english and y_left >= y_right:
                y_left = y_current - needed
            else:
                y_right = y_current - needed

            idx += 1

        page_text = f"Page {page_no}"
        if total_pages:
            page_text += f" / {total_pages}"
        c.setFont(FONT_NAME, char_size_pt)
        c.drawCentredString(page_w_pt / 2, bottom_y / 2, page_text)

        if idx < total:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# 헤더
# -----------------------
def draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, font_size_pt):
    header_y = page_h_pt - 6 * mm
    title = f"{filename_label} - 영어 단어 시험지" if filename_label else "영어 단어 시험지"
    c.setFont(FONT_NAME, font_size_pt * 2.5)
    c.drawCentredString(page_w_pt / 2, header_y, title)
    c.setLineWidth(0.5)
    c.line(12 * mm, header_y - 4 * mm, page_w_pt - 12 * mm, header_y - 4 * mm)

    meta_y = header_y - 10 * mm
    c.setFont(FONT_NAME, font_size_pt)
    left_x = page_w_pt / 2 + 10 * mm
    c.drawString(left_x, meta_y, "이름: _______________________")
    c.drawString(left_x, meta_y - 6 * mm, "학년/반: __ / __     번호: ____")
    c.drawString(left_x, meta_y - 12 * mm, "점수: ______ / 100")
    c.drawString(left_x, meta_y - 18 * mm, "시험일: ____________________")
    c.setFont(FONT_NAME, font_size_pt)

# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(layout="wide", page_title="영어 단어 시험지 생성기")
st.title("영어 단어 시험지 생성기")
st.write("영어는 왼쪽, 한글은 오른쪽 열 / 페이지당 최대한 배치 / 정확한 페이지 번호")

uploaded_files = st.file_uploader("파일 업로드 (.xlsx, .csv)", type=["xlsx", "csv"], accept_multiple_files=True)
num_questions = st.number_input("문항 수", min_value=2, max_value=500, value=60, step=2)

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        try:
            df = pd.read_excel(f) if f.name.endswith(".xlsx") else pd.read_csv(io.BytesIO(f.read()), encoding="utf-8-sig")
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
        combined = pd.concat(dfs).dropna().drop_duplicates(subset=["english"])
        pick_n = min(num_questions, len(combined))
        sampled = combined.sample(frac=1, random_state=42).iloc[:pick_n]

        # Day 라벨
        day_labels = sorted(set(extract_day_label(f.name) for f in uploaded_files),
                           key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999)
        file_label = " - ".join(day_labels) if day_labels else "영어 단어 시험지"

        # word_pairs
        half = pick_n // 2
        word_pairs = [(str(sampled.iloc[i]["english"]), str(sampled.iloc[i]["korean"]), True) for i in range(half)]
        word_pairs += [(str(sampled.iloc[i]["english"]), str(sampled.iloc[i]["korean"]), False) for i in range(half, pick_n)]

        total_pages = simulate_page_count(word_pairs, pick_n)

        test_buf = create_test_pdf(word_pairs, pick_n, file_label, total_pages)
        answer_buf = create_answer_pdf(word_pairs, pick_n, file_label, total_pages)

        st.download_button("시험지 다운로드", test_buf, "시험지.pdf", "application/pdf")
        st.download_button("정답지 다운로드", answer_buf, "정답지.pdf", "application/pdf")
        st.success(f"완료! {pick_n}문항, {total_pages}페이지")
else:
    st.info("파일을 업로드하세요.")

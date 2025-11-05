# app.py
import streamlit as st
import pandas as pd
import io
import random
import tempfile
import re
from math import ceil
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
# 폰트 안전 로딩 (한글 + 영어 지원, 자동 대체)
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
            # NotoSansKR 폰트 자동 다운로드 (로컬 전용)
            os.makedirs("fonts", exist_ok=True)
            font_path = "fonts/NotoSansKR-Regular.ttf"
            if not os.path.exists(font_path):
                st.info("NotoSansKR 폰트 다운로드 중...")
                import urllib.request
                url = "https://github.com/google/fonts/raw/main/ofl/notosanskr/NotoSansKR-Regular.ttf"
                urllib.request.urlretrieve(url, font_path)
            pdfmetrics.registerFont(TTFont("NotoSansKR", font_path))
            FONT_NAME = "NotoSansKR"
            st.success("폰트: NotoSansKR 사용 (한글+영어 완벽 지원)")
        except Exception as e:
            FONT_NAME = "Helvetica"
            st.warning("한글 폰트를 로드할 수 없습니다. 영문 폰트(Helvetica)로 대체합니다. 한글은 깨질 수 있습니다.")
            # Helvetica는 기본 등록됨

# 앱 시작 시 폰트 등록
register_korean_font()

# -----------------------
# 사용자 지정 상수 (mm 단위)
# -----------------------
NUM_X1_MM = 24       # 왼쪽 열 번호 위치
UNDER_X1_MM = 62     # 왼쪽 밑줄 시작
UNDER_LEN_MM = 37    # 밑줄 길이

NUM_X2_MM = 111      # 오른쪽 열 번호 위치
UNDER_X2_MM = 152    # 오른쪽 밑줄 시작

TOP_OFFSET_MM = 62
BOTTOM_RESERVED_MM = 24
LINE_HEIGHT_EN_MM = 4
LINE_HEIGHT_KO_GAP_MM = 4
CHAR_SIZE_MM = 2
HEADER_LEFT_MARGIN_MM = 24
HEADER_RIGHT_MARGIN_MM = 24

# -----------------------
# stringWidth 캐싱 (성능 최적화)
# -----------------------
@lru_cache(maxsize=10000)
def cached_string_width(text, font_name, font_size_pt):
    return pdfmetrics.stringWidth(text, font_name, font_size_pt)

# -----------------------
# 텍스트 래핑 유틸 (폭 기준, 캐시 사용)
# -----------------------
def wrap_text_by_width(text, font_name, font_size_pt, max_width_pt):
    if not text or not text.strip():
        return []
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
# 파일명에서 Day n 추출
# -----------------------
def extract_day_label(filename):
    name = filename.rsplit("/", 1)[-1]
    name_wo_ext = ".".join(name.split(".")[:-1]) if "." in name else name
    m = re.search(r"Day[-_ ]*(\d+)", name_wo_ext, flags=re.IGNORECASE)
    if m:
        return f"Day {int(m.group(1))}"
    return name_wo_ext

# -----------------------
# 페이지 수 시뮬레이션
# -----------------------
def simulate_page_count(word_pairs, num_questions):
    total = min(num_questions, len(word_pairs))
    if total == 0:
        return 1

    page_w_pt, page_h_pt = A4
    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm
    per_line_h = char_size_pt * 1.1

    avail_h = page_h_pt - top_offset - bottom_reserved
    y_left = page_h_pt - top_offset
    y_right = page_h_pt - top_offset
    page_count = 1
    left_used = False

    for i in range(total):
        eng, kor, is_kor_blank = word_pairs[i]
        shown_text = eng if is_kor_blank else kor
        is_english_mode = is_kor_blank

        # 열 선택
        if is_english_mode and not left_used:
            y_current = y_left
            text_max_w = (UNDER_X1_MM * mm - 2 * mm) - (NUM_X1_MM * mm + 6 * mm)
        elif not is_english_mode and y_right >= y_left:
            y_current = y_right
            text_max_w = (UNDER_X2_MM * mm - 2 * mm) - (NUM_X2_MM * mm + 6 * mm)
        else:
            page_count += 1
            y_left = page_h_pt - top_offset
            y_right = page_h_pt - top_offset
            left_used = False
            continue

        lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
        block_h = len(lines or [""]) * per_line_h
        gap_after = line_height_en if is_kor_blank else ko_gap
        needed = block_h + gap_after

        if y_current - needed < bottom_reserved:
            if (is_english_mode and left_used) or (not is_english_mode and y_right < y_left):
                page_count += 1
                y_left = page_h_pt - top_offset
                y_right = page_h_pt - top_offset
                left_used = False
            continue

        if is_english_mode and not left_used:
            y_left = y_current - needed
            left_used = True
        else:
            y_right = y_current - needed

    return page_count

# -----------------------
# PDF: 시험지 생성 (열 전환 + 페이지 번호)
# -----------------------
def create_test_pdf(word_pairs, num_questions, filename_label=None, total_pages=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    num_x1 = NUM_X1_MM * mm
    under_x1 = UNDER_X1_MM * mm
    under_len = UNDER_LEN_MM * mm
    num_x2 = NUM_X2_MM * mm
    under_x2 = UNDER_X2_MM * mm

    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    text_x1 = num_x1 + 6 * mm
    text_x2 = num_x2 + 6 * mm
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, char_size_pt)

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1

    while idx < total:
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_left = page_h_pt - top_offset
        y_right = page_h_pt - top_offset
        left_used = False

        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            is_english_mode = is_kor_blank

            # 열 선택
            if is_english_mode and not left_used:
                y_current = y_left
                num_x = num_x1
                text_x = text_x1
                under_x = under_x1
                text_max_w = text_max_w1
                column = 'left'
            elif not is_english_mode and y_right >= y_left:
                y_current = y_right
                num_x = num_x2
                text_x = text_x2
                under_x = under_x2
                text_max_w = text_max_w2
                column = 'right'
            else:
                break

            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
            if not lines:
                lines = [""]
            per_line_h = char_size_pt * 1.1
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed_space = block_h + gap_after

            if y_current - needed_space < bottom_reserved:
                break

            c.setFillColor(colors.black)
            c.drawString(num_x, y_current, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_current - li * per_line_h
                c.drawString(text_x, line_y, line)
            underline_y = y_current - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.5)
            c.line(under_x, underline_y, under_x + under_len, underline_y)

            if column == 'left':
                y_left = y_current - needed_space
                left_used = True
            else:
                y_right = y_current - needed_space

            idx += 1

        page_text = f"Page {page_no}"
        if total_pages:
            page_text += f" / {total_pages}"
        c.setFont(FONT_NAME, char_size_pt)
        c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)

        if idx < total:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# PDF: 정답지 생성
# -----------------------
def create_answer_pdf(word_pairs, num_questions, filename_label=None, total_pages=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    num_x1 = NUM_X1_MM * mm
    under_x1 = UNDER_X1_MM * mm
    under_len = UNDER_LEN_MM * mm
    num_x2 = NUM_X2_MM * mm
    under_x2 = UNDER_X2_MM * mm

    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    text_x1 = num_x1 + 6 * mm
    text_x2 = num_x2 + 6 * mm
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, char_size_pt)

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1

    while idx < total:
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_left = page_h_pt - top_offset
        y_right = page_h_pt - top_offset
        left_used = False

        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            is_english_mode = is_kor_blank

            if is_english_mode and not left_used:
                y_current = y_left
                num_x = num_x1
                text_x = text_x1
                under_x = under_x1
                text_max_w = text_max_w1
                column = 'left'
            elif not is_english_mode and y_right >= y_left:
                y_current = y_right
                num_x = num_x2
                text_x = text_x2
                under_x = under_x2
                text_max_w = text_max_w2
                column = 'right'
            else:
                break

            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
            if not lines:
                lines = [""]
            per_line_h = char_size_pt * 1.1
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed_space = block_h + gap_after

            if y_current - needed_space < bottom_reserved:
                break

            c.setFillColor(colors.black)
            c.drawString(num_x, y_current, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_current - li * per_line_h
                c.drawString(text_x, line_y, line)
            underline_y = y_current - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.line(under_x, underline_y, under_x + under_len, underline_y)

            # 정답 출력 (파란색)
            c.setFillColor(colors.blue)
            answer_text = kor if is_kor_blank else eng
            answer_lines = wrap_text_by_width(answer_text, FONT_NAME, char_size_pt, under_len - 2*mm)
            for li, a_line in enumerate(answer_lines):
                a_y = y_current - li * per_line_h
                c.drawString(under_x + 1*mm, a_y, a_line)
            c.setFillColor(colors.black)

            if column == 'left':
                y_left = y_current - needed_space
                left_used = True
            else:
                y_right = y_current - needed_space

            idx += 1

        page_text = f"Page {page_no}"
        if total_pages:
            page_text += f" / {total_pages}"
        c.setFont(FONT_NAME, char_size_pt)
        c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)

        if idx < total:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# 헤더 그리기
# -----------------------
def draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, font_size_pt):
    top_margin = 6 * mm
    header_y = page_h_pt - top_margin
    title_font_size = font_size_pt * 2.5
    small_font = font_size_pt

    title = f"{filename_label} - 영어 단어 시험지" if filename_label else "영어 단어 시험지"
    c.setFont(FONT_NAME, title_font_size)
    c.setFillColor(colors.black)
    c.drawCentredString(page_w_pt / 2, header_y, title)

    sep_y = header_y - (4 * mm)
    c.setLineWidth(0.5)
    c.line(12 * mm, sep_y, page_w_pt - 12 * mm, sep_y)

    meta_x = page_w_pt - 12 * mm
    meta_y = sep_y - (6 * mm)
    c.setFont(FONT_NAME, small_font)
    left_meta_x = page_w_pt / 2 + 10 * mm
    c.drawString(left_meta_x, meta_y, "이름: _______________________")
    c.drawString(left_meta_x + 0, meta_y - (6 * mm), "학년/반: __ / __     번호: ____")
    c.drawString(left_meta_x + 0, meta_y - (12 * mm), "점수: ______ / 100")
    c.drawString(left_meta_x + 0, meta_y - (18 * mm), "시험일: ____________________")

    c.setFont(FONT_NAME, font_size_pt)

# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(layout="wide", page_title="영어 단어 시험지 생성기")
st.title("정밀 레이아웃 영어 단어 시험지 생성기")
st.write("여러 파일 업로드 → Day 1 - Day 3 헤더, 영어는 왼쪽 열, 한글은 오른쪽 열, 페이지 번호 정확 출력")

uploaded_files = st.file_uploader("파일 업로드 (.xlsx 또는 .csv, 여러 개 가능)", type=["xlsx", "csv"], accept_multiple_files=True)
num_questions = st.number_input("출력할 전체 문항 수", min_value=2, max_value=500, value=60, step=2)

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        try:
            if str(f.name).lower().endswith(".xlsx"):
                df = pd.read_excel(f)
            else:
                raw = f.read()
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig")
                except:
                    df = pd.read_csv(io.BytesIO(raw), encoding="cp949")
            cols = [c.strip().lower() for c in df.columns]
            df.columns = cols
            if "english" in df.columns and "korean" in df.columns:
                df_sub = df[["english", "korean"]].copy()
            elif "단어" in df.columns and "뜻" in df.columns:
                df_sub = df[["단어", "뜻"]].copy()
                df_sub.columns = ["english", "korean"]
            else:
                df_sub = df.iloc[:, :2].copy()
                df_sub.columns = ["english", "korean"]
            dfs.append(df_sub)
        except Exception as e:
            st.error(f"{f.name} 처리중 오류: {e}")

    if not dfs:
        st.warning("유효한 데이터가 없습니다.")
    else:
        combined = pd.concat(dfs, ignore_index=True)
        combined = combined.dropna(subset=["english", "korean"])
        combined = combined.drop_duplicates(subset=["english"])
        available = len(combined)
        if available == 0:
            st.warning("단어가 없습니다.")
        else:
            pick_n = min(int(num_questions), available)
            sampled = combined.sample(frac=1, random_state=42).reset_index(drop=True).iloc[:pick_n]

            # Day 라벨 연결
            day_labels = []
            for f in uploaded_files:
                label = extract_day_label(f.name)
                if label and label not in day_labels:
                    day_labels.append(label)
            def day_key(x):
                m = re.search(r'Day\s*(\d+)', x, re.IGNORECASE)
                return int(m.group(1)) if m else 999
            day_labels.sort(key=day_key)
            file_label = " - ".join(day_labels) if day_labels else "영어 단어 시험지"

            # word_pairs 생성
            half = pick_n // 2
            word_pairs = []
            for i in range(half):
                eng = str(sampled.iloc[i]["english"])
                kor = str(sampled.iloc[i]["korean"])
                word_pairs.append((eng, kor, True))
            for i in range(half, pick_n):
                eng = str(sampled.iloc[i]["english"])
                kor = str(sampled.iloc[i]["korean"])
                word_pairs.append((eng, kor, False))

            # 페이지 수 계산
            total_pages = simulate_page_count(word_pairs, pick_n)

            # PDF 생성
            test_buf = create_test_pdf(word_pairs, pick_n, filename_label=file_label, total_pages=total_pages)
            answer_buf = create_answer_pdf(word_pairs, pick_n, filename_label=file_label, total_pages=total_pages)

            st.download_button("시험지 다운로드 (PDF)", data=test_buf, file_name="시험지.pdf", mime="application/pdf")
            st.download_button("정답지 다운로드 (PDF)", data=answer_buf, file_name="정답지.pdf", mime="application/pdf")

            st.success(f"총 {pick_n}문항 시험지 생성 완료 ({total_pages}페이지, 원본 단어 수: {available})")
else:
    st.info("엑셀(.xlsx) 또는 CSV 파일을 업로드해주세요.")

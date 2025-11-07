# app.py
# 영어 단어 시험지 생성기 - 정밀 레이아웃 + 랜덤 섞기 + 전체 페이지 수 표시
# + 영어 ↔ 한글 전환 시 열 자동 교체

import streamlit as st
import pandas as pd
import io
import re
from math import ceil
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import mm
from reportlab.lib import colors

# -----------------------
# 사용자 지정 상수 (mm 단위)
# -----------------------
NUM_X1_MM = 24          # 왼쪽 열 번호 위치 (mm)
UNDER_X1_MM = 62        # 왼쪽 밑줄 시작 위치 (mm)
UNDER_LEN_MM = 37       # 밑줄 길이 (mm)

NUM_X2_MM = 111         # 오른쪽 열 번호 위치 (mm)
UNDER_X2_MM = 152       # 오른쪽 밑줄 시작 위치 (mm)

TOP_OFFSET_MM = 62      # 문항 시작 위치 (상단 헤더 포함, mm)
BOTTOM_RESERVED_MM = 24 # 페이지 하단 여백 (페이지 번호 영역, mm)
LINE_HEIGHT_EN_MM = 4   # 영어 문항 간격 (mm)
LINE_HEIGHT_KO_GAP_MM = 4  # 한글 문항 간격 (mm)
CHAR_SIZE_MM = 2        # 글자 크기 (mm → pt 변환됨)
# -----------------------

# -----------------------
# 폰트 등록 (한글 지원)
# -----------------------
pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
FONT_NAME = "HYSMyeongJo-Medium"
# -----------------------

# -----------------------
# 텍스트 줄바꿈 유틸 (너비 기준)
# -----------------------
def wrap_text_by_width(text, font_name, font_size_pt, max_width_pt):
    """주어진 너비(pt)에 맞춰 텍스트 줄바꿈 (영어: 단어 단위, 한글: 문자 단위)"""
    if not text:
        return []
    text = str(text).strip()
    if not text:
        return []
    
    words = text.split(" ")
    lines = []
    cur = ""
    
    for w in words:
        candidate = (cur + " " + w).strip() if cur else w
        if pdfmetrics.stringWidth(candidate, font_name, font_size_pt) <= max_width_pt:
            cur = candidate
        else:
            if cur:
                lines.append(cur)
            if pdfmetrics.stringWidth(w, font_name, font_size_pt) <= max_width_pt:
                cur = w
            else:
                sub = ""
                for ch in w:
                    if pdfmetrics.stringWidth(sub + ch, font_name, font_size_pt) <= max_width_pt:
                        sub += ch
                    else:
                        if sub:
                            lines.append(sub)
                        sub = ch
                cur = sub
    if cur:
        lines.append(cur)
    return lines
# -----------------------

# -----------------------
# 파일명에서 Day n 추출
# -----------------------
def extract_day_label(filename):
    """파일명에서 'Day 1', 'day_2' 등을 추출 → 'Day 1' 형식으로 통일"""
    name = filename.rsplit("/", 1)[-1]
    name_wo_ext = ".".join(name.split(".")[:-1]) if "." in name else name
    m = re.search(r"(Day[-_ ]*(\d+))", name_wo_ext, flags=re.IGNORECASE)
    if m:
        return f"Day {int(m.group(2))}"
    return name_wo_ext
# -----------------------

# -----------------------
# 페이지 수 계산 시뮬레이션 (렌더링 없이)
# -----------------------
def simulate_page_count(word_pairs, num_questions):
    """PDF 생성 전에 총 페이지 수 계산 (정확한 Page X / Y 표시용)"""
    total = min(num_questions, len(word_pairs))
    if total == 0:
        return 1

    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    per_line_h = char_size_pt * 1.1
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    text_max_w1 = (UNDER_X1_MM * mm - 2 * mm) - (NUM_X1_MM * mm + 6 * mm)
    text_max_w2 = (UNDER_X2_MM * mm - 2 * mm) - (NUM_X2_MM * mm + 6 * mm)

    page_h_pt = A4[1]
    avail_h = page_h_pt - top_offset - bottom_reserved

    idx = 0
    page_count = 1
    y_left = page_h_pt - top_offset
    y_right = page_h_pt - top_offset
    cur_col = "left"                     # 현재 쓰는 열
    last_type = None                     # 이전 문항 타입 (True:영어, False:한글)

    while idx < total:
        eng, kor, is_kor_blank = word_pairs[idx]
        shown_text = eng if is_kor_blank else kor
        text_max_w = text_max_w1 if cur_col == "left" else text_max_w2
        lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
        if not lines:
            lines = [""]
        block_h = len(lines) * per_line_h
        gap_after = line_height_en if is_kor_blank else ko_gap
        needed = block_h + gap_after

        # ---- 타입 전환 시 열 교체 ----
        if last_type is not None and last_type != is_kor_blank:
            # 타입 바뀌면 반대 열로 이동
            cur_col = "right" if cur_col == "left" else "left"
            # 현재 열에 공간이 부족하면 새 페이지
            if (cur_col == "left" and y_left - needed < bottom_reserved) or \
               (cur_col == "right" and y_right - needed < bottom_reserved):
                page_count += 1
                y_left = page_h_pt - top_offset
                y_right = page_h_pt - top_offset

        # ---- 현재 열에 넣기 ----
        if cur_col == "left":
            if y_left - needed >= bottom_reserved:
                y_left -= needed
            else:
                cur_col = "right"
                if y_right - needed >= bottom_reserved:
                    y_right -= needed
                else:
                    page_count += 1
                    y_left = page_h_pt - top_offset
                    y_right = page_h_pt - top_offset
                    y_left -= needed
        else:   # right
            if y_right - needed >= bottom_reserved:
                y_right -= needed
            else:
                cur_col = "left"
                if y_left - needed >= bottom_reserved:
                    y_left -= needed
                else:
                    page_count += 1
                    y_left = page_h_pt - top_offset
                    y_right = page_h_pt - top_offset
                    y_right -= needed

        last_type = is_kor_blank
        idx += 1

    return page_count
# -----------------------

# -----------------------
# 공통: 한 열 그리기 헬퍼 (시험지·정답지 공용)
# -----------------------
def draw_column(
    c, word_pairs, start_idx, total, is_test,
    num_x, under_x, text_x, text_max_w,
    y_pos, bottom_reserved, char_size_pt,
    line_height_en, ko_gap, per_line_h
):
    """
    하나의 열(왼쪽·오른쪽)을 그린다.
    반환값: (다음 인덱스, 사용한 y 위치)
    """
    idx = start_idx
    y = y_pos
    while idx < total:
        eng, kor, is_kor_blank = word_pairs[idx]
        shown_text = eng if is_kor_blank else kor
        lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
        if not lines:
            lines = [""]
        block_h = len(lines) * per_line_h
        gap_after = line_height_en if is_kor_blank else ko_gap
        needed = block_h + gap_after

        if y - needed < bottom_reserved:
            break

        # 번호
        c.setFillColor(colors.black)
        c.drawString(num_x, y, f"{idx + 1}.")
        # 문제 텍스트
        for li, line in enumerate(lines):
            c.drawString(text_x, y - li * per_line_h, line)
        # 밑줄
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.5)
        c.line(under_x, y - 0.1 * mm, under_x + UNDER_LEN_MM * mm, y - 0.1 * mm)

        # ---- 정답지일 경우 정답 표시 ----
        if not is_test:
            c.setFillColor(colors.blue)
            answer_text = kor if is_kor_blank else eng
            answer_lines = wrap_text_by_width(answer_text, FONT_NAME, char_size_pt,
                                            UNDER_LEN_MM * mm - 2 * mm)
            for li, a_line in enumerate(answer_lines):
                c.drawString(under_x + 1 * mm, y - li * per_line_h, a_line)
            c.setFillColor(colors.black)

        y -= needed
        idx += 1
    return idx, y
# -----------------------

# -----------------------
# PDF: 시험지 생성
# -----------------------
def create_test_pdf(word_pairs, num_questions, filename_label=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    # mm → pt
    num_x1 = NUM_X1_MM * mm
    under_x1 = UNDER_X1_MM * mm
    num_x2 = NUM_X2_MM * mm
    under_x2 = UNDER_X2_MM * mm

    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm
    per_line_h = char_size_pt * 1.1

    text_x1 = num_x1 + 6 * mm
    text_x2 = num_x2 + 6 * mm
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, char_size_pt)

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1
    cur_col = "left"          # 현재 쓰는 열
    last_type = None          # 이전 문항 타입

    while idx < total:
        # 첫 페이지만 헤더
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_left = page_h_pt - top_offset
        y_right = page_h_pt - top_offset

        # ---- 타입 전환 감지 → 열 교체 ----
        eng, kor, is_kor_blank = word_pairs[idx]
        if last_type is not None and last_type != is_kor_blank:
            cur_col = "right" if cur_col == "left" else "left"
            # 현재 열에 공간 없으면 새 페이지
            if (cur_col == "left" and y_left - (per_line_h + line_height_en) < bottom_reserved) or \
               (cur_col == "right" and y_right - (per_line_h + line_height_en) < bottom_reserved):
                # 페이지 마무리
                total_pages = simulate_page_count(word_pairs, num_questions)
                page_text = f"Page {page_no} / {total_pages}"
                c.setFont(FONT_NAME, char_size_pt)
                c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)
                c.showPage()
                page_no += 1
                y_left = page_h_pt - top_offset
                y_right = page_h_pt - top_offset

        last_type = is_kor_blank

        # ---- 왼쪽 열 ----
        if cur_col == "left":
            idx, y_left = draw_column(
                c, word_pairs, idx, total, is_test=True,
                num_x=num_x1, under_x=under_x1, text_x=text_x1, text_max_w=text_max_w1,
                y_pos=y_left, bottom_reserved=bottom_reserved, char_size_pt=char_size_pt,
                line_height_en=line_height_en, ko_gap=ko_gap, per_line_h=per_line_h
            )
            if idx < total:
                cur_col = "right"
        # ---- 오른쪽 열 ----
        else:
            idx, y_right = draw_column(
                c, word_pairs, idx, total, is_test=True,
                num_x=num_x2, under_x=under_x2, text_x=text_x2, text_max_w=text_max_w2,
                y_pos=y_right, bottom_reserved=bottom_reserved, char_size_pt=char_size_pt,
                line_height_en=line_height_en, ko_gap=ko_gap, per_line_h=per_line_h
            )
            if idx < total:
                cur_col = "left"

        # ---- 페이지 넘김 ----
        if idx < total:
            total_pages = simulate_page_count(word_pairs, num_questions)
            page_text = f"Page {page_no} / {total_pages}"
            c.setFont(FONT_NAME, char_size_pt)
            c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)
            c.showPage()
            page_no += 1

    # 마지막 페이지 번호
    total_pages = simulate_page_count(word_pairs, num_questions)
    page_text = f"Page {page_no} / {total_pages}"
    c.setFont(FONT_NAME, char_size_pt)
    c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)

    c.save()
    buf.seek(0)
    return buf
# -----------------------

# -----------------------
# PDF: 정답지 생성 (시험지와 동일 로직, 정답만 추가)
# -----------------------
def create_answer_pdf(word_pairs, num_questions, filename_label=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    num_x1 = NUM_X1_MM * mm
    under_x1 = UNDER_X1_MM * mm
    num_x2 = NUM_X2_MM * mm
    under_x2 = UNDER_X2_MM * mm

    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm
    per_line_h = char_size_pt * 1.1

    text_x1 = num_x1 + 6 * mm
    text_x2 = num_x2 + 6 * mm
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    c.setFont(FONT_NAME, char_size_pt)

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1
    cur_col = "left"
    last_type = None

    while idx < total:
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_left = page_h_pt - top_offset
        y_right = page_h_pt - top_offset

        eng, kor, is_kor_blank = word_pairs[idx]
        if last_type is not None and last_type != is_kor_blank:
            cur_col = "right" if cur_col == "left" else "left"
            if (cur_col == "left" and y_left - (per_line_h + line_height_en) < bottom_reserved) or \
               (cur_col == "right" and y_right - (per_line_h + line_height_en) < bottom_reserved):
                total_pages = simulate_page_count(word_pairs, num_questions)
                page_text = f"Page {page_no} / {total_pages}"
                c.setFont(FONT_NAME, char_size_pt)
                c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)
                c.showPage()
                page_no += 1
                y_left = page_h_pt - top_offset
                y_right = page_h_pt - top_offset

        last_type = is_kor_blank

        if cur_col == "left":
            idx, y_left = draw_column(
                c, word_pairs, idx, total, is_test=False,
                num_x=num_x1, under_x=under_x1, text_x=text_x1, text_max_w=text_max_w1,
                y_pos=y_left, bottom_reserved=bottom_reserved, char_size_pt=char_size_pt,
                line_height_en=line_height_en, ko_gap=ko_gap, per_line_h=per_line_h
            )
            if idx < total:
                cur_col = "right"
        else:
            idx, y_right = draw_column(
                c, word_pairs, idx, total, is_test=False,
                num_x=num_x2, under_x=under_x2, text_x=text_x2, text_max_w=text_max_w2,
                y_pos=y_right, bottom_reserved=bottom_reserved, char_size_pt=char_size_pt,
                line_height_en=line_height_en, ko_gap=ko_gap, per_line_h=per_line_h
            )
            if idx < total:
                cur_col = "left"

        if idx < total:
            total_pages = simulate_page_count(word_pairs, num_questions)
            page_text = f"Page {page_no} / {total_pages}"
            c.setFont(FONT_NAME, char_size_pt)
            c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)
            c.showPage()
            page_no += 1

    total_pages = simulate_page_count(word_pairs, num_questions)
    page_text = f"Page {page_no} / {total_pages}"
    c.setFont(FONT_NAME, char_size_pt)
    c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_text)

    c.save()
    buf.seek(0)
    return buf
# -----------------------

# -----------------------
# 헤더 그리기 (첫 페이지)
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

    sep_y = header_y - 4 * mm
    c.setLineWidth(0.5)
    c.line(12 * mm, sep_y, page_w_pt - 12 * mm, sep_y)

    left_meta_x = page_w_pt / 2 + 10 * mm
    meta_y = sep_y - 6 * mm
    c.setFont(FONT_NAME, small_font)
    c.drawString(left_meta_x, meta_y, "이름: _______________________")
    c.drawString(left_meta_x, meta_y - 6 * mm, "학년/반: __ / __     번호: ____")
    c.drawString(left_meta_x, meta_y - 12 * mm, "점수: ______ / 100")
    c.drawString(left_meta_x, meta_y - 18 * mm, "시험일: ____________________")

    c.setFont(FONT_NAME, font_size_pt)
# -----------------------

# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(layout="wide", page_title="영어 단어 시험지 생성기")
st.title("정밀 레이아웃 영어 단어 시험지 생성기")
st.write("파일명을 헤더에 넣고, 점수란 포함. **매 실행마다 다른 순서**로 시험지 생성. **영어↔한글 전환 시 열 자동 교체**")

uploaded_files = st.file_uploader("파일 업로드 (.xlsx, .csv)", type=["xlsx", "csv"], accept_multiple_files=True)
num_questions = st.number_input("출력 문항 수", min_value=2, max_value=500, value=60, step=2)

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
            st.error(f"{f.name} 처리 오류: {e}")

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

            shuffled = combined.sample(frac=1).reset_index(drop=True)
            sampled = shuffled.iloc[:pick_n]

            day_labels = []
            for f in uploaded_files:
                label = extract_day_label(f.name)
                if label and label not in day_labels:
                    day_labels.append(label)
            
            def day_key(x):
                match = re.search(r'Day\s*(\d+)', x, re.IGNORECASE)
                return int(match.group(1)) if match else 999
            day_labels.sort(key=day_key)
            file_label = " - ".join(day_labels) if day_labels else "영어 단어 시험지"

            # 앞 절반 영어, 뒤 절반 한글
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

            test_buf = create_test_pdf(word_pairs, pick_n, filename_label=file_label)
            answer_buf = create_answer_pdf(word_pairs, pick_n, filename_label=file_label)

            st.download_button("시험지 다운로드 (PDF)", data=test_buf, file_name="시험지.pdf", mime="application/pdf")
            st.download_button("정답지 다운로드 (PDF)", data=answer_buf, file_name="정답지.pdf", mime="application/pdf")

            st.success(f"총 {pick_n}문항 생성 완료 (원본 단어: {available}개)")

else:
    st.info("엑셀(.xlsx) 또는 CSV 파일을 업로드해주세요.")
# -----------------------

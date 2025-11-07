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
from reportlab.lib.units import mm
from reportlab.lib import colors

# -----------------------
# 사용자 지정 상수 (요구대로 mm 단위)
# -----------------------
NUM_X1_MM = 24       # 왼쪽 열 번호 위치 (24mm)
UNDER_X1_MM = 62     # 왼쪽 밑줄 시작 위치 (62mm)
UNDER_LEN_MM = 37    # 밑줄 길이 (37mm)

NUM_X2_MM = 111      # 오른쪽 열 번호 위치 (111mm)
UNDER_X2_MM = 152    # 오른쪽 밑줄 시작 위치 (152mm)

TOP_OFFSET_MM = 62       # 위에서 62mm 지점부터 문항 시작 (상단 헤더 영역 포함)
BOTTOM_RESERVED_MM = 24  # 아래쪽 24mm 비우기(페이지 번호 영역)
LINE_HEIGHT_EN_MM = 4    # 영어가 작성되는 부분 행간 4mm
LINE_HEIGHT_KO_GAP_MM = 4 # 한글블록 끝난 뒤 다음 문항과의 간격 4mm
LINE_HEIGHT_BASE_MM = 8  # 기본 기준 (if needed)
CHAR_SIZE_MM = 2         # 한 글자당 2mm (약 6pt)
HEADER_LEFT_MARGIN_MM = 24  # 헤더내 좌측 여백 기준(24mm)
HEADER_RIGHT_MARGIN_MM = 24

# -----------------------
# 폰트 등록 (한글 지원)
# -----------------------
# ReportLab의 CID폰트 사용 (한글)
pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
FONT_NAME = "HYSMyeongJo-Medium"

# -----------------------
# 텍스트 래핑 유틸 (폭 기준)
# - 한국어: 공백이 적으면 문자 단위로 자름
# - 영어: 공백 단위로 자름
# -----------------------
def wrap_text_by_width(text, font_name, font_size_pt, max_width_pt):
    """주어진 최대 너비(pt)에 맞춰 텍스트를 줄바꿈한다.
       한국어/영어 혼합을 고려해서 단어 단위 우선, 너무 길면 문자 단위로 자름."""
    if text is None:
        return []
    text = str(text)
    if text.strip() == "":
        return []
    words = text.split(" ")
    lines = []
    cur = ""
    for w in words:
        candidate = (cur + " " + w).strip() if cur else w
        width = pdfmetrics.stringWidth(candidate, font_name, font_size_pt)
        if width <= max_width_pt:
            cur = candidate
        else:
            if cur:
                lines.append(cur)
            # now w might be too long; break it by characters
            if pdfmetrics.stringWidth(w, font_name, font_size_pt) <= max_width_pt:
                cur = w
            else:
                # split by characters
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
# 파일명에서 Day n 추출 유틸
# -----------------------
def extract_day_label(filename):
    name = filename.rsplit("/", 1)[-1]
    name_wo_ext = ".".join(name.split(".")[:-1]) if "." in name else name
    m = re.search(r"(Day[-_ ]*(\d+))", name_wo_ext, flags=re.IGNORECASE)
    if m:
        num = m.group(2)
        return f"Day {int(num)}"  # "Day 1", "Day 10" 형식 통일
    return name_wo_ext

# -----------------------
# PDF: 시험지 생성 (정밀 레이아웃)
# -----------------------
def create_test_pdf(word_pairs, num_questions, filename_label=None):
    """
    word_pairs: list of (eng, kor, is_kor_blank) where is_kor_blank==True => korean blank (student writes kor)
    filename_label: str for header display (e.g., "Day 1 - Day 1")
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    # convert mm to points
    num_x1 = NUM_X1_MM * mm
    under_x1 = UNDER_X1_MM * mm
    under_len = UNDER_LEN_MM * mm
    num_x2 = NUM_X2_MM * mm
    under_x2 = UNDER_X2_MM * mm

    top_offset = TOP_OFFSET_MM * mm
    bottom_reserved = BOTTOM_RESERVED_MM * mm
    char_size_pt = CHAR_SIZE_MM * mm  # font size in points
    line_height_en = LINE_HEIGHT_EN_MM * mm
    ko_gap = LINE_HEIGHT_KO_GAP_MM * mm

    # header area top_padding; we'll draw header for first page occupying the area above top_offset
    header_top_padding = page_h_pt - (8 * mm)  # start drawing header 8mm from top

    # text start x positions (text area starts a bit after number)
    text_x1 = num_x1 + 6 * mm
    text_x2 = num_x2 + 6 * mm

    # text max width before underline
    text_max_w1 = (under_x1 - 2 * mm) - text_x1
    text_max_w2 = (under_x2 - 2 * mm) - text_x2

    # set base font
    c.setFont(FONT_NAME, char_size_pt)

    # available height for questions = page_h - top_offset - bottom_reserved
    avail_h = page_h_pt - top_offset - bottom_reserved
    # compute how many minimal english lines fit per column using english line height (4mm)
    # but since Korean can be multi-line, we will place items dynamically until y runs out
    # initial y for first line (content start)
    content_start_y = page_h_pt - top_offset + ( (top_offset) - (page_h_pt - header_top_padding) ) # simplified
    # Better to set first content y at page_h_pt - top_offset (one line baseline)
    first_line_y = page_h_pt - top_offset

    total = min(num_questions, len(word_pairs))
    idx = 0
    page_no = 1
    # page loop: we iterate through items and place until page filled, then new page
    while idx < total:
        # Draw header on first page only (or repeated? user asked '첫 페이지 상단' so only first page)
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)
        # set font for content
        c.setFont(FONT_NAME, char_size_pt)
        # initial y positions for both columns (top of column content)
        y_col_start = page_h_pt - top_offset  # first baseline y for row 0 in each column
        # left column pointer y and right column pointer y
        y_left = y_col_start
        y_right = y_col_start

        # fill left column top-to-bottom until no space, then right column top-to-bottom
        # track how many items consumed on this page
        consumed_on_page = 0

        # left column
        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            # determine height of this item (in pts)
            # displayed_text is english if is_kor_blank True else korean
            shown_text = eng if is_kor_blank else kor
            # choose wrapping width for this column
            text_max_w = text_max_w1
            # compute wrapped lines
            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
            if not lines:
                lines = [""]
            # english shown usually one line; but we need to consider both
            # line height for each wrapped line: use char_size_pt * 1.1
            per_line_h = char_size_pt * 1.1
            # total block height = (num_lines-1)*per_line_h + per_line_h
            block_h = len(lines) * per_line_h
            # after block, requirement: if english shown (single-line typical) next item gap = 4mm;
            # for korean shown (multi-line possible) we must ensure next item starts ko_gap after last line
            gap_after = line_height_en if (is_kor_blank and len(lines) <=1 and all(' ' in ch for ch in [" "])) else ko_gap
            # However simpler: if english is shown (is_kor_blank True), gap_after = 4mm; else (korean shown) gap_after = ko_gap
            # We'll set:
            gap_after = line_height_en if is_kor_blank else ko_gap

            needed_space = block_h + gap_after
            # check if enough space in left column
            if (y_left - needed_space) < bottom_reserved:
                # not enough space -> break to right column
                break
            # place item: number at num_x1, text at text_x1, underline at under_x1
            # compute y for first line baseline = y_left
            num_x = num_x1
            text_x = text_x1
            under_x = under_x1
            # draw number
            c.setFillColor(colors.black)
            c.drawString(num_x, y_left, f"{idx+1}.")
            # draw wrapped lines
            for li, line in enumerate(lines):
                line_y = y_left - li * per_line_h
                c.drawString(text_x, line_y, line)
            # draw underline at under_x (fixed), length under_len
            underline_y = y_left - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.5)
            c.line(under_x, underline_y, under_x + under_len, underline_y)
            # advance y_left
            y_left = y_left - needed_space
            idx += 1
            consumed_on_page += 1

        # right column
        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            text_max_w = text_max_w2
            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w)
            if not lines:
                lines = [""]
            per_line_h = char_size_pt * 1.1
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed_space = block_h + gap_after
            if (y_right - needed_space) < bottom_reserved:
                # not enough space for this right-column item on this page -> page full
                break
            # place in right column
            num_x = num_x2
            text_x = text_x2
            under_x = under_x2
            c.setFillColor(colors.black)
            c.drawString(num_x, y_right, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_right - li * per_line_h
                c.drawString(text_x, line_y, line)
            underline_y = y_right - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.5)
            c.line(under_x, underline_y, under_x + under_len, underline_y)
            y_right = y_right - needed_space
            idx += 1
            consumed_on_page += 1

        # footer: page number area in bottom_reserved; font size = CHAR_SIZE_MM * mm
        c.setFont(FONT_NAME, char_size_pt)
        c.setFillColor(colors.black)
        total_pages = ceil(total / ( ( (page_h_pt - top_offset - bottom_reserved) // (char_size_pt*1.1 + ko_gap) ) * 2 ))  # approximate fallback
        # We want correct total_pages; easier: compute earlier using dynamic simulate? For simplicity, show current page_no and approximate total.
        # But better: compute total_pages exactly by simulation before rendering. For now we show page_no / ? ; we'll compute true total pages earlier.
        # We'll compute true total pages in a helper function outside; but here we had dynamic algorithm. To avoid a mismatch, we compute true_total_pages before calling this function.
        # We'll assume caller passed correct total_pages in global variable; but to be safe, compute simple as:
        # Instead, compute actual total_pages by simulating placement without drawing. We'll compute it outside and pass if needed.
        # For now, set text:
        page_count_text = f"Page {page_no}"
        c.drawCentredString(page_w_pt / 2, bottom_reserved / 2, page_count_text)

        # next page
        if idx < total:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# PDF: 정답지 생성 (밑줄 영역에 파란색 정답 출력)
# -----------------------
def create_answer_pdf(word_pairs, num_questions, filename_label=None):
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
        # draw header on first page
        if page_no == 1:
            draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, char_size_pt)

        y_col_start = page_h_pt - top_offset
        y_left = y_col_start
        y_right = y_col_start

        # left column
        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w1)
            if not lines:
                lines = [""]
            per_line_h = char_size_pt * 1.1
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed_space = block_h + gap_after
            if (y_left - needed_space) < bottom_reserved:
                break
            num_x = num_x1
            text_x = text_x1
            under_x = under_x1
            # draw number and text
            c.setFillColor(colors.black)
            c.drawString(num_x, y_left, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_left - li * per_line_h
                c.drawString(text_x, line_y, line)
            # draw underline as visual
            underline_y = y_left - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.line(under_x, underline_y, under_x + under_len, underline_y)
            # draw answer in underline area in blue
            c.setFillColor(colors.blue)
            answer_text = kor if is_kor_blank else eng
            # wrap answer to fit under_len
            answer_lines = wrap_text_by_width(answer_text, FONT_NAME, char_size_pt, under_len - 2*mm)
            for li, a_line in enumerate(answer_lines):
                a_y = y_left - li * per_line_h
                c.drawString(under_x + 1*mm, a_y, a_line)
            c.setFillColor(colors.black)
            y_left -= needed_space
            idx += 1

        # right column
        while idx < total:
            eng, kor, is_kor_blank = word_pairs[idx]
            shown_text = eng if is_kor_blank else kor
            lines = wrap_text_by_width(shown_text, FONT_NAME, char_size_pt, text_max_w2)
            if not lines:
                lines = [""]
            per_line_h = char_size_pt * 1.1
            block_h = len(lines) * per_line_h
            gap_after = line_height_en if is_kor_blank else ko_gap
            needed_space = block_h + gap_after
            if (y_right - needed_space) < bottom_reserved:
                break
            num_x = num_x2
            text_x = text_x2
            under_x = under_x2
            c.setFillColor(colors.black)
            c.drawString(num_x, y_right, f"{idx+1}.")
            for li, line in enumerate(lines):
                line_y = y_right - li * per_line_h
                c.drawString(text_x, line_y, line)
            underline_y = y_right - 0.1 * mm
            c.setStrokeColor(colors.black)
            c.line(under_x, underline_y, under_x + under_len, underline_y)
            c.setFillColor(colors.blue)
            answer_text = kor if is_kor_blank else eng
            answer_lines = wrap_text_by_width(answer_text, FONT_NAME, char_size_pt, under_len - 2*mm)
            for li, a_line in enumerate(answer_lines):
                a_y = y_right - li * per_line_h
                c.drawString(under_x + 1*mm, a_y, a_line)
            c.setFillColor(colors.black)
            y_right -= needed_space
            idx += 1

        # footer page no
        c.setFont(FONT_NAME, char_size_pt)
        c.drawCentredString(page_w_pt/2, bottom_reserved / 2, f"Page {page_no}")
        if idx < total:
            c.showPage()
            page_no += 1

    c.save()
    buf.seek(0)
    return buf

# -----------------------
# Draw header helper (first page)
# -----------------------
def draw_header_on_canvas(c, page_w_pt, page_h_pt, filename_label, font_size_pt):
    # header area between top and top_offset (TOP_OFFSET_MM)
    top_margin = 6 * mm
    header_y = page_h_pt - top_margin
    title_font_size = font_size_pt * 2.5  # make title larger but still small
    small_font = font_size_pt

    # Draw title centered (use filename_label if provided)
    title = f"{filename_label} - 영어 단어 시험지" if filename_label else "영어 단어 시험지"
    c.setFont(FONT_NAME, title_font_size)
    c.setFillColor(colors.black)
    c.drawCentredString(page_w_pt / 2, header_y, title)

    # Draw a thin separator line under title
    sep_y = header_y - (4 * mm)
    c.setLineWidth(0.5)
    c.line(12 * mm, sep_y, page_w_pt - 12 * mm, sep_y)

    # Draw metadata fields on the right side of header: 이름 / 반 / 점수 / 시험일
    meta_x = page_w_pt - 12 * mm
    meta_y = sep_y - (6 * mm)
    c.setFont(FONT_NAME, small_font)
    # Draw fields left-aligned a bit to the left from meta_x
    left_meta_x = page_w_pt / 2 + 10 * mm
    # Labels and blanks
    c.drawString(left_meta_x, meta_y, "이름: _______________________")
    c.drawString(left_meta_x + 0, meta_y - (6 * mm), "학년/반: __ / __     번호: ____")
    c.drawString(left_meta_x + 0, meta_y - (12 * mm), "점수: ______ / 100")
    c.drawString(left_meta_x + 0, meta_y - (18 * mm), "시험일: ____________________")

    # reset font
    c.setFont(FONT_NAME, font_size_pt)

# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(layout="wide", page_title="영어 단어 시험지 생성기 (정밀 레이아웃)")
st.title("📘 정밀 레이아웃 영어 단어 시험지 생성기")
st.write("파일명을 헤더에 자동으로 넣고, 점수란을 포함합니다. 한글 줄바꿈 후 다음 문항과의 간격이 4mm가 되도록 처리합니다.")

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
                except Exception:
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

            # 진짜 랜덤으로 섞기 (random_state 제거)
            shuffled = combined.sample(frac=1).reset_index(drop=True)
            sampled = shuffled.iloc[:pick_n]

            # === Day 라벨 연결 로직 (기존 유지) ===
            day_labels = []
            for f in uploaded_files:
                label = extract_day_label(f.name)
                if label and label not in day_labels:
                    day_labels.append(label)
            
            def day_key(x):
                import re
                match = re.search(r'Day\s*(\d+)', x, re.IGNORECASE)
                return int(match.group(1)) if match else 999
            day_labels.sort(key=day_key)
            
            file_label = " - ".join(day_labels) if day_labels else "영어 단어 시험지"
            # ===============================

            # Build pairs (기존 배치 로직 유지)
            half = pick_n // 2
            word_pairs = []
            for i in range(half):
                eng = str(sampled.iloc[i]["english"])
                kor = str(sampled.iloc[i]["korean"])
                word_pairs.append((eng, kor, True))   # 영어 보여주기
            for i in range(half, pick_n):
                eng = str(sampled.iloc[i]["english"])
                kor = str(sampled.iloc[i]["korean"])
                word_pairs.append((eng, kor, False))  # 한글 보여주기

            test_buf = create_test_pdf(word_pairs, pick_n, filename_label=file_label)
            answer_buf = create_answer_pdf(word_pairs, pick_n, filename_label=file_label)

            st.download_button("시험지 다운로드 (PDF)", data=test_buf, file_name="시험지.pdf", mime="application/pdf")
            st.download_button("정답지 다운로드 (PDF)", data=answer_buf, file_name="정답지.pdf", mime="application/pdf")

            st.success(f"총 {pick_n}문항으로 시험지 및 정답지 생성 완료 (원본 단어 수: {available}).")
else:
    st.info("엑셀(.xlsx) 또는 CSV 파일을 업로드해주세요.")

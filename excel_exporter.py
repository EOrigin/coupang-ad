#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
쿠팡 스마트광고 성과 분석 엑셀(XLSX) 내보내기 모듈
- 캠페인별 성과 요약 및 상세 내역 분리
- 캠페인별 개별 시트 탭 제공
- 컬럼 구성: 옵션ID, 광고비, ROAS, 매출, 상품명
"""

import os
import re
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 스타일 상수 정의
FONT_FAMILY = "맑은 고딕"

FONT_MAIN_TITLE = Font(name=FONT_FAMILY, size=15, bold=True, color="1E3A8A")
FONT_SUBTITLE = Font(name=FONT_FAMILY, size=9, color="64748B")
FONT_SECTION_TITLE = Font(name=FONT_FAMILY, size=12, bold=True, color="1E3A8A")

FONT_HEADER = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
FONT_SUBHEADER = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
FONT_DATA = Font(name=FONT_FAMILY, size=10, color="0F172A")
FONT_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="0F172A")
FONT_TOTAL = Font(name=FONT_FAMILY, size=10, bold=True, color="1E3A8A")

FILL_NAVY = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
FILL_BLUE = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
FILL_HEADER_GRAY = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
FILL_SUBTOTAL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_TOTAL_ROW = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

BORDER_THIN_COLOR = "CBD5E1"
BORDER_THIN = Border(
    left=Side(style="thin", color=BORDER_THIN_COLOR),
    right=Side(style="thin", color=BORDER_THIN_COLOR),
    top=Side(style="thin", color=BORDER_THIN_COLOR),
    bottom=Side(style="thin", color=BORDER_THIN_COLOR),
)
BORDER_HEADER = Border(
    left=Side(style="thin", color="94A3B8"),
    right=Side(style="thin", color="94A3B8"),
    top=Side(style="thin", color="94A3B8"),
    bottom=Side(style="medium", color="64748B"),
)
BORDER_SUBTOTAL = Border(
    left=Side(style="thin", color=BORDER_THIN_COLOR),
    right=Side(style="thin", color=BORDER_THIN_COLOR),
    top=Side(style="thin", color="94A3B8"),
    bottom=Side(style="thin", color="94A3B8"),
)
BORDER_TOTAL = Border(
    left=Side(style="thin", color=BORDER_THIN_COLOR),
    right=Side(style="thin", color=BORDER_THIN_COLOR),
    top=Side(style="thin", color="1E3A8A"),
    bottom=Side(style="double", color="1E3A8A"),
)

NUM_FORMAT_CURRENCY = "#,##0"
NUM_FORMAT_ROAS = '0.00"%"'
NUM_FORMAT_COUNT = '#,##0"건"'
NUM_FORMAT_TEXT = "@"


def calc_display_width(val) -> int:
    """한글/영문 혼용 문자열의 표시 너비 계산"""
    if val is None:
        return 0
    s = str(val)
    w = 0
    for ch in s:
        w += 2 if ord(ch) > 127 else 1
    return w


def auto_fit_columns(ws, min_widths=None, max_widths=None):
    """워크시트의 각 열 너비 자동 조정"""
    min_widths = min_widths or {}
    max_widths = max_widths or {}

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            # 병합된 셀이나 행 전체 타이틀 등은 너비 계산에서 과대평가 방지
            if cell.coordinate in ws.merged_cells:
                continue
            if cell.value is not None:
                val_str = str(cell.value)
                # 제목 줄 등 너무 긴 셀은 제외 (50자 이상인 경우 헤더가 아니면 제외)
                if len(val_str) > 70 and cell.row <= 3:
                    continue
                max_len = max(max_len, calc_display_width(val_str))

        calculated_width = max(max_len + 3, min_widths.get(col_letter, 10))
        if col_letter in max_widths:
            calculated_width = min(calculated_width, max_widths[col_letter])
        ws.column_dimensions[col_letter].width = calculated_width


def sanitize_sheet_title(name: str, used_names: set) -> str:
    """엑셀 시트 이름 제약 조건 준수 (최대 31자, 금지 문자 제거, 중복 방지)"""
    clean = re.sub(r'[\\/*?:\[\]]', '_', str(name)).strip()
    if not clean:
        clean = "캠페인"
    base = clean[:25]
    result = base
    counter = 1
    while result in used_names:
        result = f"{base}_{counter}"[:31]
        counter += 1
    used_names.add(result)
    return result


def export_to_excel(xlsx_path: str, all_results_by_campaign: dict, min_cost: int = 5000, max_roas: float = 600.0):
    """
    크롤링된 캠페인별 데이터를 스타일이 적용된 Excel(.xlsx) 파일로 저장
    - Sheet 1: [캠페인별 요약] 성과 요약 테이블 + 캠페인별 구분 상세 목록
    - Sheet 2: [전체 상품목록] 필터 적용 가능한 전체 플랫 데이터
    - Sheet 3+: 각 캠페인별 개별 시트 (옵션ID, 광고비, ROAS, 매출, 상품명)
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 기본 생성된 Sheet 삭제

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ==========================================
    # Sheet 1: 캠페인별 요약 (메인 대시보드 시트)
    # ==========================================
    ws_summary = wb.create_sheet(title="캠페인별 요약")
    ws_summary.views.sheetView[0].showGridLines = True

    # 1-1. 헤더 타이틀
    ws_summary.row_dimensions[1].height = 28
    ws_summary.row_dimensions[2].height = 18

    cell_title = ws_summary.cell(row=1, column=1, value="쿠팡 스마트광고 성과 분석 리포트")
    cell_title.font = FONT_MAIN_TITLE

    cell_sub = ws_summary.cell(
        row=2,
        column=1,
        value=f"수집일시: {now_str}  |  기준: 집행 광고비 {min_cost:,}원 이상, ROAS {max_roas:.1f}% 미만"
    )
    cell_sub.font = FONT_SUBTITLE

    # 1-2. [섹션 1] 캠페인별 성과 요약 테이블
    ws_summary.row_dimensions[4].height = 22
    cell_sec1 = ws_summary.cell(row=4, column=1, value="■ 1. 캠페인별 성과 요약")
    cell_sec1.font = FONT_SECTION_TITLE

    summary_headers = ["캠페인명", "조건부합 상품수", "총 광고비", "총 매출", "평균 ROAS"]
    ws_summary.row_dimensions[5].height = 24
    for c_idx, h in enumerate(summary_headers, start=1):
        c = ws_summary.cell(row=5, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_NAVY
        c.alignment = ALIGN_CENTER
        c.border = BORDER_HEADER

    cur_row = 6
    total_all_items = 0
    total_all_cost = 0
    total_all_sales = 0

    for cname, items in all_results_by_campaign.items():
        c_cost = sum(it.get("ad_cost", 0) for it in items)
        c_sales = sum(it.get("sales", 0) for it in items)
        c_count = len(items)
        c_roas = (c_sales / c_cost * 100.0) if c_cost > 0 else 0.0

        total_all_items += c_count
        total_all_cost += c_cost
        total_all_sales += c_sales

        ws_summary.row_dimensions[cur_row].height = 20

        # 캠페인명
        c1 = ws_summary.cell(row=cur_row, column=1, value=cname)
        c1.font = FONT_DATA
        c1.alignment = ALIGN_LEFT
        c1.border = BORDER_THIN

        # 상품수
        c2 = ws_summary.cell(row=cur_row, column=2, value=c_count)
        c2.font = FONT_DATA
        c2.alignment = ALIGN_CENTER
        c2.number_format = NUM_FORMAT_COUNT
        c2.border = BORDER_THIN

        # 총 광고비
        c3 = ws_summary.cell(row=cur_row, column=3, value=c_cost)
        c3.font = FONT_DATA
        c3.alignment = ALIGN_RIGHT
        c3.number_format = NUM_FORMAT_CURRENCY
        c3.border = BORDER_THIN

        # 총 매출
        c4 = ws_summary.cell(row=cur_row, column=4, value=c_sales)
        c4.font = FONT_DATA
        c4.alignment = ALIGN_RIGHT
        c4.number_format = NUM_FORMAT_CURRENCY
        c4.border = BORDER_THIN

        # 평균 ROAS
        c5 = ws_summary.cell(row=cur_row, column=5, value=c_roas)
        c5.font = FONT_DATA
        c5.alignment = ALIGN_RIGHT
        c5.number_format = NUM_FORMAT_ROAS
        c5.border = BORDER_THIN

        cur_row += 1

    # 요약 합계 행
    overall_roas = (total_all_sales / total_all_cost * 100.0) if total_all_cost > 0 else 0.0
    ws_summary.row_dimensions[cur_row].height = 22

    t1 = ws_summary.cell(row=cur_row, column=1, value="[전체 합계]")
    t1.font = FONT_TOTAL
    t1.alignment = ALIGN_CENTER
    t1.fill = FILL_TOTAL_ROW
    t1.border = BORDER_TOTAL

    t2 = ws_summary.cell(row=cur_row, column=2, value=total_all_items)
    t2.font = FONT_TOTAL
    t2.alignment = ALIGN_CENTER
    t2.number_format = NUM_FORMAT_COUNT
    t2.fill = FILL_TOTAL_ROW
    t2.border = BORDER_TOTAL

    t3 = ws_summary.cell(row=cur_row, column=3, value=total_all_cost)
    t3.font = FONT_TOTAL
    t3.alignment = ALIGN_RIGHT
    t3.number_format = NUM_FORMAT_CURRENCY
    t3.fill = FILL_TOTAL_ROW
    t3.border = BORDER_TOTAL

    t4 = ws_summary.cell(row=cur_row, column=4, value=total_all_sales)
    t4.font = FONT_TOTAL
    t4.alignment = ALIGN_RIGHT
    t4.number_format = NUM_FORMAT_CURRENCY
    t4.fill = FILL_TOTAL_ROW
    t4.border = BORDER_TOTAL

    t5 = ws_summary.cell(row=cur_row, column=5, value=overall_roas)
    t5.font = FONT_TOTAL
    t5.alignment = ALIGN_RIGHT
    t5.number_format = NUM_FORMAT_ROAS
    t5.fill = FILL_TOTAL_ROW
    t5.border = BORDER_TOTAL

    cur_row += 3  # 간격 띄우기

    # 1-3. [섹션 2] 캠페인별 상품 상세 내역 (구분형 테이블)
    ws_summary.row_dimensions[cur_row].height = 22
    cell_sec2 = ws_summary.cell(row=cur_row, column=1, value="■ 2. 캠페인별 상품 상세 내역")
    cell_sec2.font = FONT_SECTION_TITLE
    cur_row += 1

    detail_headers = ["옵션ID", "광고비", "ROAS", "매출", "상품명"]

    for cname, items in all_results_by_campaign.items():
        # 캠페인 타이틀 배너
        ws_summary.row_dimensions[cur_row].height = 24
        banner_text = f"▶ [캠페인] {cname}  (총 {len(items)}개 상품)"
        
        # A~E 열 병합 배너
        ws_summary.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=5)
        banner_cell = ws_summary.cell(row=cur_row, column=1, value=banner_text)
        banner_cell.font = FONT_HEADER
        banner_cell.fill = FILL_BLUE
        banner_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, 6):
            ws_summary.cell(row=cur_row, column=col).border = BORDER_THIN
        cur_row += 1

        # 컬럼 헤더
        ws_summary.row_dimensions[cur_row].height = 22
        for c_idx, h in enumerate(detail_headers, start=1):
            c = ws_summary.cell(row=cur_row, column=c_idx, value=h)
            c.font = FONT_SUBHEADER
            c.fill = FILL_HEADER_GRAY
            c.alignment = ALIGN_CENTER
            c.border = BORDER_HEADER
        cur_row += 1

        if not items:
            ws_summary.row_dimensions[cur_row].height = 20
            ws_summary.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=5)
            empty_cell = ws_summary.cell(row=cur_row, column=1, value="(해당 캠페인에 조건 부합 상품 없음)")
            empty_cell.font = FONT_DATA
            empty_cell.alignment = ALIGN_CENTER
            for col in range(1, 6):
                ws_summary.cell(row=cur_row, column=col).border = BORDER_THIN
            cur_row += 2
            continue

        c_cost_sum = 0
        c_sales_sum = 0

        for item in items:
            ws_summary.row_dimensions[cur_row].height = 20

            cost = item.get("ad_cost", 0)
            sales = item.get("sales", 0)
            roas = item.get("roas", 0.0)
            c_cost_sum += cost
            c_sales_sum += sales

            # 1. 옵션ID (텍스트 형태)
            d1 = ws_summary.cell(row=cur_row, column=1, value=str(item.get("option_id", "")))
            d1.font = FONT_DATA
            d1.alignment = ALIGN_CENTER
            d1.number_format = NUM_FORMAT_TEXT
            d1.border = BORDER_THIN

            # 2. 광고비
            d2 = ws_summary.cell(row=cur_row, column=2, value=cost)
            d2.font = FONT_DATA
            d2.alignment = ALIGN_RIGHT
            d2.number_format = NUM_FORMAT_CURRENCY
            d2.border = BORDER_THIN

            # 3. ROAS
            d3 = ws_summary.cell(row=cur_row, column=3, value=roas)
            d3.font = FONT_DATA
            d3.alignment = ALIGN_RIGHT
            d3.number_format = NUM_FORMAT_ROAS
            d3.border = BORDER_THIN

            # 4. 매출
            d4 = ws_summary.cell(row=cur_row, column=4, value=sales)
            d4.font = FONT_DATA
            d4.alignment = ALIGN_RIGHT
            d4.number_format = NUM_FORMAT_CURRENCY
            d4.border = BORDER_THIN

            # 5. 상품명
            d5 = ws_summary.cell(row=cur_row, column=5, value=item.get("product_name", ""))
            d5.font = FONT_DATA
            d5.alignment = ALIGN_LEFT
            d5.border = BORDER_THIN

            cur_row += 1

        # 캠페인별 소계
        camp_roas = (c_sales_sum / c_cost_sum * 100.0) if c_cost_sum > 0 else 0.0
        ws_summary.row_dimensions[cur_row].height = 22

        s1 = ws_summary.cell(row=cur_row, column=1, value=f"[{cname} 소계]")
        s1.font = FONT_BOLD
        s1.alignment = ALIGN_CENTER
        s1.fill = FILL_SUBTOTAL
        s1.border = BORDER_SUBTOTAL

        s2 = ws_summary.cell(row=cur_row, column=2, value=c_cost_sum)
        s2.font = FONT_BOLD
        s2.alignment = ALIGN_RIGHT
        s2.number_format = NUM_FORMAT_CURRENCY
        s2.fill = FILL_SUBTOTAL
        s2.border = BORDER_SUBTOTAL

        s3 = ws_summary.cell(row=cur_row, column=3, value=camp_roas)
        s3.font = FONT_BOLD
        s3.alignment = ALIGN_RIGHT
        s3.number_format = NUM_FORMAT_ROAS
        s3.fill = FILL_SUBTOTAL
        s3.border = BORDER_SUBTOTAL

        s4 = ws_summary.cell(row=cur_row, column=4, value=c_sales_sum)
        s4.font = FONT_BOLD
        s4.alignment = ALIGN_RIGHT
        s4.number_format = NUM_FORMAT_CURRENCY
        s4.fill = FILL_SUBTOTAL
        s4.border = BORDER_SUBTOTAL

        s5 = ws_summary.cell(row=cur_row, column=5, value=f"총 {len(items)}개 상품")
        s5.font = FONT_BOLD
        s5.alignment = ALIGN_LEFT
        s5.fill = FILL_SUBTOTAL
        s5.border = BORDER_SUBTOTAL

        cur_row += 3  # 다음 캠페인과의 간격

    auto_fit_columns(
        ws_summary,
        min_widths={"A": 18, "B": 15, "C": 14, "D": 15, "E": 45},
        max_widths={"E": 70}
    )

    # ==========================================
    # Sheet 2: 전체 상품목록 (전체 플랫 테이블 & 필터)
    # ==========================================
    ws_all = wb.create_sheet(title="전체_상품목록")
    ws_all.views.sheetView[0].showGridLines = True

    all_headers = ["캠페인명", "옵션ID", "광고비", "ROAS", "매출", "상품명"]
    ws_all.row_dimensions[1].height = 24
    for c_idx, h in enumerate(all_headers, start=1):
        c = ws_all.cell(row=1, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_NAVY
        c.alignment = ALIGN_CENTER
        c.border = BORDER_HEADER

    row_idx = 2
    for cname, items in all_results_by_campaign.items():
        for item in items:
            ws_all.row_dimensions[row_idx].height = 20

            # 캠페인명
            c1 = ws_all.cell(row=row_idx, column=1, value=cname)
            c1.font = FONT_DATA
            c1.alignment = ALIGN_LEFT
            c1.border = BORDER_THIN

            # 옵션ID
            c2 = ws_all.cell(row=row_idx, column=2, value=str(item.get("option_id", "")))
            c2.font = FONT_DATA
            c2.alignment = ALIGN_CENTER
            c2.number_format = NUM_FORMAT_TEXT
            c2.border = BORDER_THIN

            # 광고비
            c3 = ws_all.cell(row=row_idx, column=3, value=item.get("ad_cost", 0))
            c3.font = FONT_DATA
            c3.alignment = ALIGN_RIGHT
            c3.number_format = NUM_FORMAT_CURRENCY
            c3.border = BORDER_THIN

            # ROAS
            c4 = ws_all.cell(row=row_idx, column=4, value=item.get("roas", 0.0))
            c4.font = FONT_DATA
            c4.alignment = ALIGN_RIGHT
            c4.number_format = NUM_FORMAT_ROAS
            c4.border = BORDER_THIN

            # 매출
            c5 = ws_all.cell(row=row_idx, column=5, value=item.get("sales", 0))
            c5.font = FONT_DATA
            c5.alignment = ALIGN_RIGHT
            c5.number_format = NUM_FORMAT_CURRENCY
            c5.border = BORDER_THIN

            # 상품명
            c6 = ws_all.cell(row=row_idx, column=6, value=item.get("product_name", ""))
            c6.font = FONT_DATA
            c6.alignment = ALIGN_LEFT
            c6.border = BORDER_THIN

            row_idx += 1

    # 전체 시트 자동 필터 및 틀 고정
    if row_idx > 2:
        ws_all.auto_filter.ref = f"A1:F{row_idx - 1}"
    ws_all.freeze_panes = "A2"
    auto_fit_columns(
        ws_all,
        min_widths={"A": 16, "B": 16, "C": 14, "D": 12, "E": 14, "F": 45},
        max_widths={"F": 70}
    )

    # ==========================================
    # Sheet 3+: 캠페인별 개별 시트 탭
    # ==========================================
    used_sheet_names = {"캠페인별 요약", "전체_상품목록"}

    for cname, items in all_results_by_campaign.items():
        if not items:
            continue  # 조건 부합 상품이 없는 캠페인은 별도 탭 생성 생략

        sheet_title = sanitize_sheet_title(cname, used_sheet_names)
        ws_camp = wb.create_sheet(title=sheet_title)
        ws_camp.views.sheetView[0].showGridLines = True

        # 캠페인 타이틀
        ws_camp.row_dimensions[1].height = 24
        t_cell = ws_camp.cell(row=1, column=1, value=f"[캠페인] {cname}")
        t_cell.font = FONT_MAIN_TITLE

        # 컬럼 헤더 (요청된 컬럼: 옵션ID, 광고비, ROAS, 매출, 상품명)
        camp_headers = ["옵션ID", "광고비", "ROAS", "매출", "상품명"]
        ws_camp.row_dimensions[3].height = 24
        for c_idx, h in enumerate(camp_headers, start=1):
            c = ws_camp.cell(row=3, column=c_idx, value=h)
            c.font = FONT_HEADER
            c.fill = FILL_BLUE
            c.alignment = ALIGN_CENTER
            c.border = BORDER_HEADER

        c_row = 4
        c_cost_total = 0
        c_sales_total = 0

        for item in items:
            ws_camp.row_dimensions[c_row].height = 20

            cost = item.get("ad_cost", 0)
            sales = item.get("sales", 0)
            roas = item.get("roas", 0.0)
            c_cost_total += cost
            c_sales_total += sales

            # 옵션ID
            c1 = ws_camp.cell(row=c_row, column=1, value=str(item.get("option_id", "")))
            c1.font = FONT_DATA
            c1.alignment = ALIGN_CENTER
            c1.number_format = NUM_FORMAT_TEXT
            c1.border = BORDER_THIN

            # 광고비
            c2 = ws_camp.cell(row=c_row, column=2, value=cost)
            c2.font = FONT_DATA
            c2.alignment = ALIGN_RIGHT
            c2.number_format = NUM_FORMAT_CURRENCY
            c2.border = BORDER_THIN

            # ROAS
            c3 = ws_camp.cell(row=c_row, column=3, value=roas)
            c3.font = FONT_DATA
            c3.alignment = ALIGN_RIGHT
            c3.number_format = NUM_FORMAT_ROAS
            c3.border = BORDER_THIN

            # 매출
            c4 = ws_camp.cell(row=c_row, column=4, value=sales)
            c4.font = FONT_DATA
            c4.alignment = ALIGN_RIGHT
            c4.number_format = NUM_FORMAT_CURRENCY
            c4.border = BORDER_THIN

            # 상품명
            c5 = ws_camp.cell(row=c_row, column=5, value=item.get("product_name", ""))
            c5.font = FONT_DATA
            c5.alignment = ALIGN_LEFT
            c5.border = BORDER_THIN

            c_row += 1

        # 소계 행
        c_camp_roas = (c_sales_total / c_cost_total * 100.0) if c_cost_total > 0 else 0.0
        ws_camp.row_dimensions[c_row].height = 22

        s1 = ws_camp.cell(row=c_row, column=1, value="[합계]")
        s1.font = FONT_BOLD
        s1.alignment = ALIGN_CENTER
        s1.fill = FILL_SUBTOTAL
        s1.border = BORDER_TOTAL

        s2 = ws_camp.cell(row=c_row, column=2, value=c_cost_total)
        s2.font = FONT_BOLD
        s2.alignment = ALIGN_RIGHT
        s2.number_format = NUM_FORMAT_CURRENCY
        s2.fill = FILL_SUBTOTAL
        s2.border = BORDER_TOTAL

        s3 = ws_camp.cell(row=c_row, column=3, value=c_camp_roas)
        s3.font = FONT_BOLD
        s3.alignment = ALIGN_RIGHT
        s3.number_format = NUM_FORMAT_ROAS
        s3.fill = FILL_SUBTOTAL
        s3.border = BORDER_TOTAL

        s4 = ws_camp.cell(row=c_row, column=4, value=c_sales_total)
        s4.font = FONT_BOLD
        s4.alignment = ALIGN_RIGHT
        s4.number_format = NUM_FORMAT_CURRENCY
        s4.fill = FILL_SUBTOTAL
        s4.border = BORDER_TOTAL

        s5 = ws_camp.cell(row=c_row, column=5, value=f"총 {len(items)}개 상품")
        s5.font = FONT_BOLD
        s5.alignment = ALIGN_LEFT
        s5.fill = FILL_SUBTOTAL
        s5.border = BORDER_TOTAL

        if c_row > 4:
            ws_camp.auto_filter.ref = f"A3:E{c_row - 1}"
        ws_camp.freeze_panes = "A4"

        auto_fit_columns(
            ws_camp,
            min_widths={"A": 16, "B": 14, "C": 12, "D": 14, "E": 45},
            max_widths={"E": 70}
        )

    # 기본 활성 시트를 [캠페인별 요약]으로 설정
    wb.active = ws_summary

    wb.save(xlsx_path)
    return xlsx_path


if __name__ == "__main__":
    import sys
    json_input = sys.argv[1] if len(sys.argv) > 1 else "output/coupang_ads_20261012_145541.json"
    xlsx_output = sys.argv[2] if len(sys.argv) > 2 else json_input.replace(".json", ".xlsx")

    with open(json_input, "r", encoding="utf-8") as f:
        data = json.load(f)

    out = export_to_excel(xlsx_output, data)
    print(f"XLSX 저장 성공: {out}")

"""10명 제출 내용을 원본 HWPX 양식과 유사한 표 구조의 PDF로 렌더링.

레이아웃:
 - 제목
 - 사업단 공통확인사항 (실적/계획 2칸)
 - 메인 본문 표: [과제|분야|이름|구분|업무실적|업무계획]
 - 스마트돌봄스페이스, 회의자료 3종 (c4/c5 형식)
"""
import io
import os
import html
from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, black, blue, whitesmoke, lightgrey, grey
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from team_config import TEAM_MEMBERS, FIELD_LABELS

FONT_NAME = "KoreanFont"
_BUNDLED_FONT = Path(__file__).parent / "fonts" / "NanumGothic-Regular.ttf"
_FONT_CANDIDATES = [
    str(_BUNDLED_FONT),
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "C:\\Windows\\Fonts\\malgun.ttf",
    "C:\\Windows\\Fonts\\NanumGothic.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
]


def _register_font() -> str:
    for path in _FONT_CANDIDATES:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(FONT_NAME, path))
                return FONT_NAME
            except Exception:
                continue
    return "Helvetica"


def _escape(s: str) -> str:
    return html.escape(s or "")


def _mk_body_paragraph(text: str, style: ParagraphStyle,
                       acquired_prefix: str = "") -> Paragraph:
    """한 셀 내용을 Paragraph로 변환. 줄바꿈은 <br/>.
    acquired_prefix 있으면 파란색으로 맨 위에 추가."""
    lines = (text or "").split("\n")
    body_html = "<br/>".join(_escape(l) for l in lines if l is not None)
    if acquired_prefix:
        prefix_lines = acquired_prefix.split("\n")
        prefix_html = "<br/>".join(_escape(l) for l in prefix_lines if l is not None)
        full_html = f'<font color="#0000FF">{prefix_html}</font>'
        if body_html:
            full_html += f'<br/>{body_html}'
        return Paragraph(full_html, style)
    return Paragraph(body_html, style)


def build_pdf(submissions: dict, title_date: str,
              period_start: str, period_end: str,
              plan_start: str, plan_end: str) -> bytes:
    font = _register_font()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=8*mm, rightMargin=8*mm,
        topMargin=10*mm, bottomMargin=10*mm,
        title=f"돌봄로봇 주간 업무보고 ({title_date})",
    )

    # 스타일 정의
    title_style = ParagraphStyle(
        "title", fontName=font, fontSize=14, leading=18,
        alignment=1, spaceAfter=6, textColor=black,
    )
    header_cell = ParagraphStyle(
        "header_cell", fontName=font, fontSize=9, leading=12,
        alignment=1, textColor=black,
    )
    label_cell = ParagraphStyle(
        "label_cell", fontName=font, fontSize=9, leading=12,
        alignment=1, textColor=black,
    )
    body_cell = ParagraphStyle(
        "body_cell", fontName=font, fontSize=8.5, leading=11,
        alignment=0, textColor=black,
    )

    story = []
    story.append(Paragraph(f"과업별 업무 보고 ({title_date})", title_style))

    # ─── 1) 사업단 공통확인사항 박스 ─────────────────────────
    jjs = submissions.get("정지수", {})
    chm = submissions.get("최혜민", {})
    pc_header = [
        Paragraph("<b>사업단 공통확인사항</b>", header_cell),
        Paragraph(f"<b>업무 실적 ({period_start} ~ {period_end})</b>", header_cell),
        Paragraph(f"<b>업무 계획 ({plan_start} ~ {plan_end})</b>", header_cell),
    ]
    pc_rows = [
        pc_header,
        ["", _mk_body_paragraph(jjs.get("project_confirmation_1", ""), body_cell),
             _mk_body_paragraph(chm.get("project_confirmation_2_plan", ""), body_cell)],
    ]
    pc_tbl = Table(pc_rows, colWidths=[40*mm, 120*mm, 120*mm])
    pc_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("GRID", (0, 0), (-1, -1), 0.5, grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (-1, 0), whitesmoke),
        ("BACKGROUND", (0, 1), (0, 1), whitesmoke),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("VALIGN", (0, 1), (0, 1), "MIDDLE"),
        ("SPAN", (0, 0), (0, 1)),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(pc_tbl)
    story.append(Spacer(1, 4*mm))

    # ─── 2) 메인 본문 표 ─────────────────────────────────
    # 컬럼: 과제 | 분야 | 이름 | 구분 | 실적 | 계획
    col_widths = [20*mm, 22*mm, 18*mm, 14*mm, 113*mm, 93*mm]

    # 헤더 행
    main_rows = [[
        Paragraph("<b>구분</b>", header_cell), "", "", "",
        Paragraph(f"<b>업무 실적<br/>({period_start} ~ {period_end})</b>", header_cell),
        Paragraph(f"<b>업무 계획<br/>({plan_start} ~ {plan_end})</b>", header_cell),
    ]]

    # SPAN 및 배경 스타일 모음
    styles = [
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("GRID", (0, 0), (-1, -1), 0.4, grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (-1, 0), whitesmoke),
        ("ALIGN", (0, 0), (3, 0), "CENTER"),
        ("ALIGN", (4, 0), (5, 0), "CENTER"),
        ("SPAN", (0, 0), (3, 0)),  # 헤더 "구분"이 0~3열 차지
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]

    row_idx = 1
    for m in TEAM_MEMBERS:
        data = submissions.get(m["name"], {})
        cat1 = m.get("category1", "")
        cat2 = m.get("category2", "")

        if m["has_research"]:
            # 2행 구조: 연구 / 업무
            # 연구 행 — 획득데이터(파란) + 연구실적
            acquired = data.get("acquired_data", "")
            research_done_cell = _mk_body_paragraph(
                data.get("research_done", ""), body_cell,
                acquired_prefix=f"획득 데이터: {acquired}" if acquired else ""
            )
            research_plan_cell = _mk_body_paragraph(
                data.get("research_plan", ""), body_cell)
            main_rows.append([
                Paragraph(f"<b>{cat1}</b>", label_cell),
                Paragraph(cat2, label_cell),
                Paragraph(f"<b>{m['name']}</b>", label_cell),
                Paragraph("연구", label_cell),
                research_done_cell, research_plan_cell,
            ])
            # 업무 행
            main_rows.append([
                "", "", "", Paragraph("업무", label_cell),
                _mk_body_paragraph(data.get("task_done", ""), body_cell),
                _mk_body_paragraph(data.get("task_plan", ""), body_cell),
            ])
            # SPAN: 과제/분야/이름 2행 병합
            styles.append(("SPAN", (0, row_idx), (0, row_idx + 1)))
            styles.append(("SPAN", (1, row_idx), (1, row_idx + 1)))
            styles.append(("SPAN", (2, row_idx), (2, row_idx + 1)))
            styles.append(("ALIGN", (0, row_idx), (3, row_idx + 1), "CENTER"))
            styles.append(("VALIGN", (0, row_idx), (3, row_idx + 1), "MIDDLE"))
            row_idx += 2
        else:
            # 1행 구조: 업무만
            main_rows.append([
                Paragraph(f"<b>{cat1}</b>", label_cell),
                Paragraph(cat2, label_cell),
                Paragraph(f"<b>{m['name']}</b>", label_cell),
                Paragraph("업무", label_cell),
                _mk_body_paragraph(data.get("task_done", ""), body_cell),
                _mk_body_paragraph(data.get("task_plan", ""), body_cell),
            ])
            styles.append(("ALIGN", (0, row_idx), (3, row_idx), "CENTER"))
            styles.append(("VALIGN", (0, row_idx), (3, row_idx), "MIDDLE"))
            row_idx += 1

    main_tbl = Table(main_rows, colWidths=col_widths, repeatRows=1)
    main_tbl.setStyle(TableStyle(styles))
    story.append(main_tbl)
    story.append(Spacer(1, 4*mm))

    # ─── 3) 하단 특별 항목: 스마트돌봄스페이스 / 회의자료 3종 ────
    bjs = submissions.get("백정은", {})
    chm = submissions.get("최혜민", {})
    bottom_rows = [
        [Paragraph("<b>구분</b>", header_cell),
         Paragraph("<b>업무 실적</b>", header_cell),
         Paragraph("<b>업무 계획</b>", header_cell)],
        [Paragraph("<b>스마트돌봄스페이스</b><br/><font size=7>(백정은)</font>", label_cell),
         _mk_body_paragraph(bjs.get("smart_care_space_done", ""), body_cell),
         _mk_body_paragraph(bjs.get("smart_care_space_plan", ""), body_cell)],
        [Paragraph("<b>연구소 회의자료</b><br/><font size=7>(소장주재회의)</font>", label_cell),
         _mk_body_paragraph(chm.get("research_meeting", ""), body_cell), ""],
        [Paragraph("<b>원장+재활원<br/>주요간부회의자료</b>", label_cell),
         _mk_body_paragraph(chm.get("director_meeting", ""), body_cell), ""],
        [Paragraph("<b>복지부 본부<br/>주간일정 보산진</b>", label_cell),
         _mk_body_paragraph(chm.get("mohw_weekly", ""), body_cell), ""],
    ]
    bottom_tbl = Table(bottom_rows, colWidths=[40*mm, 120*mm, 120*mm])
    bottom_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("GRID", (0, 0), (-1, -1), 0.4, grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (-1, 0), whitesmoke),
        ("BACKGROUND", (0, 1), (0, -1), whitesmoke),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("VALIGN", (0, 1), (0, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(bottom_tbl)

    doc.build(story)
    return buf.getvalue()

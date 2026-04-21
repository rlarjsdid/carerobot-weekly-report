"""10명 제출 내용을 PDF로 렌더링 (한글 폰트 자동 등록).

Streamlit Cloud(Linux): /usr/share/fonts/truetype/nanum/NanumGothic.ttf
Windows 로컬: C:\\Windows\\Fonts\\malgun.ttf
macOS: /System/Library/Fonts/AppleSDGothicNeo.ttc
"""
import io
import os
from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, black, blue, whitesmoke, lightgrey
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from team_config import TEAM_MEMBERS, FIELD_LABELS, get_fields_for

FONT_NAME = "KoreanFont"
# 번들 폰트가 최우선 (배포 환경 독립적으로 확실히 작동)
_BUNDLED_FONT = Path(__file__).parent / "fonts" / "NanumGothic-Regular.ttf"
_FONT_CANDIDATES = [
    str(_BUNDLED_FONT),
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "C:\\Windows\\Fonts\\malgun.ttf",
    "C:\\Windows\\Fonts\\NanumGothic.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    "/Library/Fonts/NanumGothic.ttf",
]


def _register_font() -> str:
    for path in _FONT_CANDIDATES:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(FONT_NAME, path))
                return FONT_NAME
            except Exception:
                continue
    return "Helvetica"  # 최후 폴백 (한글 깨짐)


def build_pdf(submissions: dict, title_date: str,
              period_start: str, period_end: str,
              plan_start: str, plan_end: str) -> bytes:
    """submissions = {이름: {필드키: 텍스트}}"""
    font = _register_font()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=12*mm, bottomMargin=12*mm,
        title=f"돌봄로봇 주간 업무보고 ({title_date})",
    )

    title_style = ParagraphStyle(
        "title", fontName=font, fontSize=16, leading=22,
        alignment=1, spaceAfter=10, textColor=black,
    )
    header_style = ParagraphStyle(
        "header", fontName=font, fontSize=11, leading=16,
        alignment=1, textColor=black,
    )
    name_style = ParagraphStyle(
        "name", fontName=font, fontSize=13, leading=18,
        textColor=black, spaceBefore=8, spaceAfter=4,
    )
    label_style = ParagraphStyle(
        "label", fontName=font, fontSize=9, leading=12,
        textColor=HexColor("#555555"), spaceBefore=2,
    )
    body_style = ParagraphStyle(
        "body", fontName=font, fontSize=9.5, leading=13,
        textColor=black, leftIndent=6,
    )
    blue_body_style = ParagraphStyle(
        "bluebody", fontName=font, fontSize=9.5, leading=13,
        textColor=blue, leftIndent=6,
    )

    story = []
    story.append(Paragraph(
        f"과업별 업무 보고 ({title_date})", title_style))
    story.append(Paragraph(
        f"실적: {period_start} ~ {period_end} | 계획: {plan_start} ~ {plan_end}",
        header_style,
    ))
    story.append(Spacer(1, 6*mm))

    BLUE_FIELDS = {"acquired_data"}

    for m in TEAM_MEMBERS:
        data = submissions.get(m["name"], {})
        fields = get_fields_for(m)
        if not any(data.get(f) for f in fields):
            # 제출 없는 사람은 skip
            continue

        # 이름 + 분류
        cat = f"[{m['category1']} - {m['category2']}]" if m["category2"] else f"[{m['category1']}]"
        story.append(Paragraph(
            f"<b>{m['name']}</b> &nbsp;&nbsp; <font size=9 color='#888'>{cat}</font>",
            name_style,
        ))

        # 각 필드
        rows = []
        for f in fields:
            val = (data.get(f) or "").strip()
            if not val:
                continue
            style = blue_body_style if f in BLUE_FIELDS else body_style
            lines = val.split("\n")
            body_html = "<br/>".join(_escape_for_rl(l) for l in lines)
            rows.append([
                Paragraph(f"<b>{FIELD_LABELS[f]}</b>", label_style),
                Paragraph(body_html, style),
            ])

        if rows:
            tbl = Table(rows, colWidths=[50*mm, 220*mm])
            tbl.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOX", (0, 0), (-1, -1), 0.4, lightgrey),
                ("INNERGRID", (0, 0), (-1, -1), 0.2, lightgrey),
                ("BACKGROUND", (0, 0), (0, -1), whitesmoke),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))
            story.append(tbl)

    doc.build(story)
    return buf.getvalue()


def _escape_for_rl(s: str) -> str:
    """ReportLab Paragraph은 <b>, <i>, <br/> 같은 태그를 해석하므로
    일반 텍스트의 특수문자는 escape."""
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

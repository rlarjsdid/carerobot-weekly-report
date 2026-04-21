"""구글시트를 저장소로 사용하는 모듈.

시트 구조:
  A: 이름  B: 주차  C~K: 10개 필드  L: 제출시간
"""
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import streamlit as st
from team_config import MEMBER_NAMES

KST = timezone(timedelta(hours=9))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

FIELD_KEYS = [
    "acquired_data",
    "research_done",
    "research_plan",
    "task_done",
    "task_plan",
    "smart_care_space_done",
    "smart_care_space_plan",
    "project_confirmation_1",
    "project_confirmation_2_done",
    "project_confirmation_2_plan",
    "research_meeting",
    "director_meeting",
    "mohw_weekly",
]

FIELD_LABELS_KR = {
    "acquired_data": "획득데이터",
    "research_done": "연구실적",
    "research_plan": "연구계획",
    "task_done": "업무실적",
    "task_plan": "업무계획",
    "smart_care_space_done": "스마트돌봄스페이스_실적",
    "smart_care_space_plan": "스마트돌봄스페이스_계획",
    "project_confirmation_1": "사업단공통확인사항1",
    "project_confirmation_2_done": "사업단공통확인사항2_실적",
    "project_confirmation_2_plan": "사업단공통확인사항2_계획",
    "research_meeting": "연구소회의자료",
    "director_meeting": "주요간부회의자료",
    "mohw_weekly": "보산진주간일정",
}

HEADER = ["이름", "주차"] + [FIELD_LABELS_KR[k] for k in FIELD_KEYS] + ["제출시간"]
COL_COUNT = len(HEADER)


@st.cache_resource
def _get_client():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def _get_sheet():
    client = _get_client()
    sheet_id = st.secrets["sheet"]["id"]
    ss = client.open_by_key(sheet_id)
    try:
        ws = ss.worksheet("submissions")
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title="submissions", rows=500, cols=COL_COUNT + 2)
        ws.append_row(HEADER)
    return ws


def _ensure_header(ws):
    values = ws.row_values(1)
    if values != HEADER:
        end_col = chr(ord('A') + COL_COUNT - 1)
        ws.update(f"A1:{end_col}1", [HEADER])


def _row_to_dict(row: list) -> dict:
    padded = list(row) + [""] * (COL_COUNT - len(row))
    return dict(zip(HEADER, padded))


def load_week(week: str) -> dict:
    ws = _get_sheet()
    _ensure_header(ws)
    records = ws.get_all_values()
    out = {}
    for r in records[1:]:
        row = _row_to_dict(r)
        if row.get("주차", "") == week:
            data = {k: row.get(FIELD_LABELS_KR[k], "") for k in FIELD_KEYS}
            data["submitted_at"] = row.get("제출시간", "")
            out[row["이름"]] = data
    return out


def save_submission(name: str, week: str, values: dict) -> str:
    """values = {필드키: 텍스트} — FIELD_KEYS의 부분 집합. 누락은 빈 문자열로 저장."""
    ws = _get_sheet()
    _ensure_header(ws)
    now = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    field_values = [values.get(k, "") for k in FIELD_KEYS]
    new_row = [name, week] + field_values + [now]

    all_rows = ws.get_all_values()
    for i, r in enumerate(all_rows[1:], start=2):
        row = _row_to_dict(r)
        if row.get("이름") == name and row.get("주차", "") == week:
            end_col = chr(ord('A') + COL_COUNT - 1)
            ws.update(f"A{i}:{end_col}{i}", [new_row])
            return "updated"
    ws.append_row(new_row)
    return "created"


def submission_status(week: str) -> list:
    data = load_week(week)
    out = []
    for name in MEMBER_NAMES:
        r = data.get(name)
        out.append({
            "name": name,
            "submitted": r is not None,
            "submitted_at": r["submitted_at"] if r else "",
        })
    return out

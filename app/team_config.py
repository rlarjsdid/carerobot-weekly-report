"""팀원 정보 및 HWPX 셀 매핑.

각 필드는 다음 형식의 튜플로 지정:
    (col, row)                         → 기본 검정색, 첫 번째 매치
    (col, row, color)                  → 지정 색 ("black"/"blue"), 첫 번째 매치
    (col, row, color, nth)             → nth번째 매치 (0-indexed) — 같은 주소가 여러 테이블에 있을 때

표에서 공유되는 고정 필드는 SHARED_FIELDS에 정의.
"""

# 고정 4필드: 본인 영역의 연구/업무 실적+계획
# (has_research=False 인 사람은 research_* 필드 무시)

TEAM_MEMBERS = [
    {
        "name": "백정은", "category1": "본부과제", "category2": "현장실증",
        "has_research": True,
        "cells": {
            # 획득 데이터 (파란색, 현장실증팀 전용)
            "acquired_data": (4, 1, "blue"),
            "research_done": (4, 2),
            "research_plan": (5, 2),
            "task_done": (4, 3),
            "task_plan": (5, 3),
            # 스마트돌봄스페이스 (백정은 전용) — 실적/계획 2칸
            "smart_care_space_done": (4, 27),
            "smart_care_space_plan": (5, 27),
        },
    },
    {
        "name": "한벼리", "category1": "본부과제", "category2": "현장실증",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 4, "blue"),
            "research_done": (4, 5),
            "research_plan": (5, 5),
            "task_done": (4, 6),
            "task_plan": (5, 6),
        },
    },
    {
        "name": "박재우", "category1": "본부과제", "category2": "현장실증",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 7, "blue"),
            "research_done": (4, 8),
            "research_plan": (5, 8),
            "task_done": (4, 9),
            "task_plan": (5, 9),
        },
    },
    {
        "name": "이윤환", "category1": "본부과제", "category2": "현장실증",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 10, "blue"),
            "research_done": (4, 11),
            "research_plan": (5, 11),
            "task_done": (4, 12),
            "task_plan": (5, 12),
        },
    },
    {
        "name": "김건양", "category1": "본부과제", "category2": "로봇기술",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 13, "blue"),
            "research_done": (4, 14),
            "research_plan": (5, 14),
            "task_done": (4, 15),
            "task_plan": (5, 15),
        },
    },
    {
        "name": "류현경", "category1": "본부과제", "category2": "로봇기술",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 16, "blue"),
            "research_done": (4, 17),
            "research_plan": (5, 17),
            "task_done": (4, 18),
            "task_plan": (5, 18),
        },
    },
    {
        "name": "남재엽", "category1": "본부과제", "category2": "로봇기술",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 19, "blue"),
            "research_done": (4, 20),
            "research_plan": (5, 20),
            "task_done": (4, 21),
            "task_plan": (5, 21),
        },
    },
    {
        "name": "이경진", "category1": "본부과제", "category2": "로봇기술",
        "has_research": True,
        "cells": {
            "acquired_data": (4, 22, "blue"),
            "research_done": (4, 23),
            "research_plan": (5, 23),
            "task_done": (4, 24),
            "task_plan": (5, 24),
        },
    },
    {
        "name": "최혜민", "category1": "운영과제", "category2": "예산",
        "has_research": False,
        "cells": {
            "task_done": (4, 25),
            "task_plan": (5, 25),
            # 사업단 공통확인사항 2 - 업무실적/계획 2칸 (HWPX 매핑 보류)
            "project_confirmation_2_done": None,
            "project_confirmation_2_plan": None,
            # 회의자료 3종
            "research_meeting": (4, 28),
            "director_meeting": (4, 29),
            "mohw_weekly": (4, 30),
        },
    },
    {
        "name": "정지수", "category1": "세부과제관리", "category2": "",
        "has_research": False,
        "cells": {
            "task_done": (4, 26),
            "task_plan": (5, 26),
            # 사업단 공통확인사항 1 - 1칸 (HWPX 매핑 보류)
            "project_confirmation_1": None,
        },
    },
]

# 화면 라벨
FIELD_LABELS = {
    "acquired_data": "획득 데이터 (파란색으로 출력됨)",
    "research_done": "연구 실적",
    "research_plan": "연구 계획",
    "task_done": "업무 실적",
    "task_plan": "업무 계획",
    "smart_care_space_done": "스마트돌봄스페이스 — 업무 실적",
    "smart_care_space_plan": "스마트돌봄스페이스 — 업무 계획",
    "project_confirmation_1": "사업단 공통확인사항 1",
    "project_confirmation_2_done": "사업단 공통확인사항 2 — 업무 실적",
    "project_confirmation_2_plan": "사업단 공통확인사항 2 — 업무 계획",
    "research_meeting": "연구소 회의자료 (소장주재회의)",
    "director_meeting": "원장+재활원 주요간부회의자료",
    "mohw_weekly": "복지부 본부 주간일정 (보산진 보고)",
}

MEMBER_NAMES = [m["name"] for m in TEAM_MEMBERS]


def get_member(name):
    for m in TEAM_MEMBERS:
        if m["name"] == name:
            return m
    return None


def get_fields_for(member):
    """해당 팀원이 작성해야 할 필드 이름 리스트 (표시 순서대로)."""
    order = ["acquired_data", "research_done", "research_plan",
             "task_done", "task_plan",
             "smart_care_space_done", "smart_care_space_plan",
             "project_confirmation_1",
             "project_confirmation_2_done", "project_confirmation_2_plan",
             "research_meeting", "director_meeting", "mohw_weekly"]
    return [f for f in order if f in member["cells"]]


APP_PASSWORD = "carerobot"
ADMIN_PASSWORD = "carerobot-admin"

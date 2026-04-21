"""돌봄로봇 주간 업무보고 취합 웹앱."""
import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path

from team_config import (
    TEAM_MEMBERS, MEMBER_NAMES, FIELD_LABELS,
    get_member, get_fields_for,
    APP_PASSWORD, ADMIN_PASSWORD,
)
from sheets_store import load_week, save_submission, submission_status, FIELD_KEYS
from hwpx_exporter import build_report
from pdf_exporter import build_pdf
from docx_exporter import build_docx

st.set_page_config(page_title="돌봄로봇 주간 업무보고", page_icon="📋", layout="wide")


def this_wednesday() -> str:
    today = datetime.now().date()
    days_until_wed = (2 - today.weekday()) % 7
    wednesday = today + timedelta(days=days_until_wed)
    return wednesday.strftime("%Y-%m-%d")


def wednesday_of_week(week_str: str) -> datetime:
    return datetime.strptime(week_str, "%Y-%m-%d")


def auth_gate():
    # URL 쿼리 파라미터로 로그인 유지 (새로고침 대응)
    qp = st.query_params
    if not st.session_state.get("authed"):
        token = qp.get("auth")
        if token == "team":
            st.session_state["authed"] = True
            st.session_state["is_admin"] = False
        elif token == "admin":
            st.session_state["authed"] = True
            st.session_state["is_admin"] = True

    if st.session_state.get("authed"):
        return True

    st.title("📋 돌봄로봇 주간 업무보고")
    pw = st.text_input("비밀번호", type="password", key="pw_input")
    if st.button("입장"):
        if pw == APP_PASSWORD:
            st.session_state["authed"] = True
            st.session_state["is_admin"] = False
            st.query_params["auth"] = "team"
            st.rerun()
        elif pw == ADMIN_PASSWORD:
            st.session_state["authed"] = True
            st.session_state["is_admin"] = True
            st.query_params["auth"] = "admin"
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False


def member_page():
    st.header("✍️ 업무보고 작성")

    col1, col2 = st.columns([2, 2])
    with col1:
        name = st.selectbox("본인 이름", MEMBER_NAMES, key="member_name")
    with col2:
        week = st.text_input("보고 주차 (수요일 기준)", value=this_wednesday(),
                             help="예: 2026-04-22")

    member = get_member(name)
    fields = get_fields_for(member)

    current = load_week(week).get(name, {})

    # 지난주 제출 내용 조회 (이번주 초기값으로 사용)
    last_week = None
    last_week_data = {}
    try:
        this_wed = wednesday_of_week(week)
        last_week = (this_wed - timedelta(days=7)).strftime("%Y-%m-%d")
        last_week_data = load_week(last_week).get(name, {})
    except Exception:
        pass

    # prefill 우선순위: 이번주 기존 저장본 > 지난주 내용 > 빈값
    if current:
        existing = current
        st.info(f"📝 이번주({week}) 저장본을 불러왔습니다. (마지막 저장: {current.get('submitted_at','-')})")
    elif last_week_data:
        existing = last_week_data
        st.warning(f"🗂️ **지난주({last_week}) 내용을 그대로 불러왔습니다.** 내용을 확인하고 이번주에 맞게 수정해주세요.")
    else:
        existing = {}
        st.caption(f"ℹ️ 지난주({last_week or '-'}) 제출 기록도 없어 빈 칸으로 시작합니다.")

    values = {}
    with st.form("report_form", clear_on_submit=False):
        if "acquired_data" in fields:
            st.subheader("📊 획득 데이터")
            st.caption("입력한 내용은 최종 보고서에 **파란색**으로 출력됩니다.")
            values["acquired_data"] = st.text_area(
                FIELD_LABELS["acquired_data"],
                value=existing.get("acquired_data", ""),
                height=120,
                placeholder="예: Obi + 진동센서, 미니스위치 데이터(○○○ 가정실증)",
                label_visibility="collapsed",
            )

        if member["has_research"]:
            st.subheader("🔬 연구")
            rc1, rc2 = st.columns(2)
            with rc1:
                values["research_done"] = st.text_area(
                    FIELD_LABELS["research_done"],
                    value=existing.get("research_done", ""),
                    height=220, placeholder="한 줄에 한 항목씩 작성",
                )
            with rc2:
                values["research_plan"] = st.text_area(
                    FIELD_LABELS["research_plan"],
                    value=existing.get("research_plan", ""),
                    height=220, placeholder="한 줄에 한 항목씩 작성",
                )

        if "task_done" in fields:
            st.subheader("📝 업무")
            tc1, tc2 = st.columns(2)
            with tc1:
                values["task_done"] = st.text_area(
                    FIELD_LABELS["task_done"],
                    value=existing.get("task_done", ""),
                    height=220, placeholder="한 줄에 한 항목씩 작성",
                )
            with tc2:
                values["task_plan"] = st.text_area(
                    FIELD_LABELS["task_plan"],
                    value=existing.get("task_plan", ""),
                    height=220, placeholder="한 줄에 한 항목씩 작성",
                )

        extra_fields = [f for f in fields if f in (
            "smart_care_space_done", "smart_care_space_plan",
            "project_confirmation_1",
            "project_confirmation_2_done", "project_confirmation_2_plan",
            "research_meeting", "director_meeting", "mohw_weekly")]
        if extra_fields:
            st.subheader("📌 추가 작성 항목")
            for f in extra_fields:
                values[f] = st.text_area(
                    FIELD_LABELS[f],
                    value=existing.get(f, ""),
                    height=150,
                    key=f"extra_{f}",
                )

        submitted = st.form_submit_button("💾 저장 / 제출", use_container_width=True)

    if submitted:
        try:
            action = save_submission(name, week, values)
            st.success(f"저장 완료 ({'신규 제출' if action=='created' else '기존 내용 수정'})")
            # 개인 백업 텍스트 생성 → 다운로드 버튼 제공
            lines = [f"=== {name} / {week} ===\n"]
            for f in get_fields_for(member):
                v = values.get(f, "") or ""
                lines.append(f"\n[{FIELD_LABELS[f]}]\n{v}\n")
            backup_txt = "".join(lines).encode('utf-8')
            st.download_button(
                "📄 내 제출본 TXT 백업 다운로드 (권장: 매주 저장해두세요)",
                data=backup_txt,
                file_name=f"{name}_{week}.txt",
                mime="text/plain",
            )
        except Exception as e:
            st.error(f"저장 실패: {e}")


def admin_page():
    st.header("📊 담당자 대시보드")

    week = st.text_input("조회 주차", value=this_wednesday(),
                         help="예: 2026-04-22 (해당 주 수요일)")

    status = submission_status(week)
    df = pd.DataFrame([
        {"이름": s["name"],
         "상태": "✅ 완료" if s["submitted"] else "⏳ 미제출",
         "제출시간": s["submitted_at"] or "-"}
        for s in status
    ])

    done_count = sum(1 for s in status if s["submitted"])
    st.metric("제출 현황", f"{done_count} / {len(status)}")
    st.dataframe(df, use_container_width=True, hide_index=True)

    missing = [s["name"] for s in status if not s["submitted"]]
    if missing:
        st.warning(f"미제출: {', '.join(missing)}")
    else:
        st.success("전원 제출 완료 🎉")

    with st.expander("🔍 제출 내용 미리보기"):
        data = load_week(week)
        for name in MEMBER_NAMES:
            r = data.get(name)
            if not r:
                continue
            st.markdown(f"**{name}**  _{r['submitted_at']}_")
            member = get_member(name)
            fields = get_fields_for(member)
            for f in fields:
                val = r.get(f, "") or "-"
                st.caption(FIELD_LABELS[f])
                st.text(val)
            st.divider()

    st.subheader("📤 내보내기")

    try:
        wed = wednesday_of_week(week)
    except ValueError:
        st.error("주차 형식이 잘못되었습니다 (YYYY-MM-DD).")
        return

    # Word(.docx) 다운로드 섹션 — 기본 출력 (HWPX 대체)
    st.markdown("### 📘 Word(.docx) — 회의자료 (권장)")
    st.caption("한컴오피스·워드 모두에서 열림. 원본 HWPX 양식과 동일 구조.")

    if st.button("📘 Word 생성 및 다운로드", type="primary", use_container_width=True):
        try:
            subs_for_docx = load_week(week)
            last_week_str = (wed - timedelta(days=7)).strftime("%Y-%m-%d")
            last_week_subs = load_week(last_week_str)
            # 필드 단위 fallback: 이번주 비어있는 필드만 지난주 내용 사용
            for name in MEMBER_NAMES:
                cur = subs_for_docx.get(name, {})
                last = last_week_subs.get(name, {})
                merged = dict(last)  # 지난주 값 기본
                for k, v in cur.items():
                    if v:  # 이번주 비어있지 않으면 덮어씀
                        merged[k] = v
                if merged:
                    subs_for_docx[name] = merged

            docx_bytes = build_docx(
                subs_for_docx,
                title_date=wed.strftime("%y.%m.%d."),
                period_start=(wed - timedelta(days=7)).strftime("%Y.%m.%d."),
                period_end=(wed - timedelta(days=1)).strftime("%Y.%m.%d."),
                plan_start=wed.strftime("%Y.%m.%d."),
                plan_end=(wed + timedelta(days=6)).strftime("%Y.%m.%d."),
            )
            st.download_button(
                "💾 Word 다운로드",
                data=docx_bytes,
                file_name=f"돌봄로봇_업무보고({wed.strftime('%m.%d')})_취합본.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.success("Word 파일 생성 완료. 한컴오피스 한글이나 MS Word에서 열립니다.")
        except Exception as e:
            st.error(f"Word 생성 실패: {e}")

    st.markdown("---")

    # PDF 다운로드 섹션 — 백업
    st.markdown("### 📕 PDF (백업)")
    st.caption("⚠️ 레이아웃 일부 이슈 있음. Word 우선 권장.")

    if st.button("📕 PDF 생성 및 다운로드", use_container_width=True):
        try:
            subs_for_pdf = load_week(week)
            # 미제출자 지난주 fallback 동일 적용
            last_week_str = (wed - timedelta(days=7)).strftime("%Y-%m-%d")
            last_week_subs = load_week(last_week_str)
            for name in MEMBER_NAMES:
                if name not in subs_for_pdf and name in last_week_subs:
                    subs_for_pdf[name] = last_week_subs[name]

            pdf_bytes = build_pdf(
                subs_for_pdf,
                title_date=wed.strftime("%y.%m.%d."),
                period_start=(wed - timedelta(days=7)).strftime("%Y.%m.%d."),
                period_end=(wed - timedelta(days=1)).strftime("%Y.%m.%d."),
                plan_start=wed.strftime("%Y.%m.%d."),
                plan_end=(wed + timedelta(days=6)).strftime("%Y.%m.%d."),
            )
            st.download_button(
                "💾 PDF 다운로드",
                data=pdf_bytes,
                file_name=f"돌봄로봇_업무보고({wed.strftime('%m.%d')})_취합본.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            st.success("PDF 생성 완료.")
        except Exception as e:
            st.error(f"PDF 생성 실패: {e}")

    st.markdown("---")
    st.markdown("### 📄 HWPX (한글 편집용)")
    st.caption("⚠️ 현재 한글에서 안 열리는 이슈 조사 중.")

    # 🔬 디버그: 템플릿 원본 그대로 다운로드 (수정 없음)
    with st.expander("🔬 디버그 도구 (HWPX 이슈 진단용)"):
        st.caption(
            "선택한 템플릿 파일을 **내용 수정 없이 그대로** 다운로드. "
            "한글에서 이 파일이 열리면 → XML 수정 로직이 문제. "
            "이것도 안 열리면 → Streamlit 다운로드 경로나 ZIP 재포장이 문제."
        )
        repo_root = Path(__file__).resolve().parent.parent
        template_files = sorted(repo_root.glob("돌봄로봇_업무보고*.hwpx"))
        if template_files:
            debug_tpl = st.selectbox(
                "원본 그대로 다운로드할 템플릿",
                template_files,
                format_func=lambda p: p.name,
                index=len(template_files) - 1,
                key="debug_tpl_select",
            )
            if st.button("🔬 원본 그대로 다운로드 (수정 0)", use_container_width=True):
                raw = debug_tpl.read_bytes()
                st.download_button(
                    "💾 원본 다운로드",
                    data=raw,
                    file_name=f"DEBUG_원본_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("다운 후 한글에서 열어보세요.")

            st.markdown("---")
            st.caption("**증분 디버그**: 어느 수정이 깨뜨리는지 좁히기")

            import zipfile
            import io as _io

            # Level 1: ZIP 재포장만 (section0.xml 내용은 원본 그대로)
            if st.button("🔬 Lv1: ZIP 재포장만 (내용 수정 0)",
                         use_container_width=True):
                from hwpx_exporter import _patch_zip_flag_bits
                src = debug_tpl.read_bytes()
                with zipfile.ZipFile(_io.BytesIO(src), 'r') as zin:
                    order = [i.filename for i in zin.infolist()]
                    infos = {i.filename: i for i in zin.infolist()}
                    files = {n: zin.read(n) for n in order}
                buf = _io.BytesIO()
                with zipfile.ZipFile(buf, 'w') as zout:
                    for name in order:
                        orig = infos[name]
                        zi = zipfile.ZipInfo(name, date_time=orig.date_time)
                        zi.compress_type = orig.compress_type
                        zi.external_attr = orig.external_attr
                        zi.create_system = orig.create_system
                        zi.create_version = orig.create_version
                        zi.extract_version = orig.extract_version
                        zi.flag_bits = orig.flag_bits
                        zi.extra = orig.extra
                        zout.writestr(zi, files[name])
                result = _patch_zip_flag_bits(buf.getvalue(), infos)
                st.download_button(
                    "💾 Lv1 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv1_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv1 = 수정 전혀 없이 ZIP만 다시 패키징. 안 열리면 ZIP 재포장 자체가 문제.")

            # Level 2: title/기간 date regex 치환만
            if st.button("🔬 Lv2: title+기간 날짜만 수정",
                         use_container_width=True):
                import re
                from hwpx_exporter import _patch_zip_flag_bits
                src = debug_tpl.read_bytes()
                with zipfile.ZipFile(_io.BytesIO(src), 'r') as zin:
                    order = [i.filename for i in zin.infolist()]
                    infos = {i.filename: i for i in zin.infolist()}
                    files = {n: zin.read(n) for n in order}
                xml = files['Contents/section0.xml'].decode('utf-8')
                xml = re.sub(r'과업별 업무 보고 \(\d{2}\.\d{2}\.\d{2}\.\)',
                             f'과업별 업무 보고 ({wed.strftime("%y.%m.%d.")})', xml)
                xml = re.sub(
                    r'업무 실적\(\d{4}\.\d{2}\.\d{2}\. ~ \d{4}\.\d{2}\.\d{2}\.\)',
                    f'업무 실적({(wed - timedelta(days=7)).strftime("%Y.%m.%d.")} ~ {(wed - timedelta(days=1)).strftime("%Y.%m.%d.")})', xml)
                xml = re.sub(
                    r'업무 계획\(\d{4}\.\d{2}\.\d{2}\. ~ \d{4}\.\d{2}\.\d{2}\.\)',
                    f'업무 계획({wed.strftime("%Y.%m.%d.")} ~ {(wed + timedelta(days=6)).strftime("%Y.%m.%d.")})', xml)
                files['Contents/section0.xml'] = xml.encode('utf-8')

                buf = _io.BytesIO()
                with zipfile.ZipFile(buf, 'w') as zout:
                    for name in order:
                        orig = infos[name]
                        zi = zipfile.ZipInfo(name, date_time=orig.date_time)
                        zi.compress_type = orig.compress_type
                        zi.external_attr = orig.external_attr
                        zi.create_system = orig.create_system
                        zi.create_version = orig.create_version
                        zi.extract_version = orig.extract_version
                        zi.flag_bits = orig.flag_bits
                        zi.extra = orig.extra
                        zout.writestr(zi, files[name])
                result = _patch_zip_flag_bits(buf.getvalue(), infos)
                st.download_button(
                    "💾 Lv2 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv2_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv2 = 제목 날짜 + 실적/계획 기간 날짜만 regex 로 치환. 셀 내용 수정은 0.")

            # Level 3: Lv2 + 김건양 연구실적 1개 셀만 수정
            if st.button("🔬 Lv3: Lv2 + 김건양 연구실적 1칸만 수정",
                         use_container_width=True):
                import re
                from hwpx_exporter import _patch_zip_flag_bits, replace_cell
                src = debug_tpl.read_bytes()
                with zipfile.ZipFile(_io.BytesIO(src), 'r') as zin:
                    order = [i.filename for i in zin.infolist()]
                    infos = {i.filename: i for i in zin.infolist()}
                    files = {n: zin.read(n) for n in order}
                xml = files['Contents/section0.xml'].decode('utf-8')
                xml = re.sub(r'과업별 업무 보고 \(\d{2}\.\d{2}\.\d{2}\.\)',
                             f'과업별 업무 보고 ({wed.strftime("%y.%m.%d.")})', xml)
                xml = re.sub(
                    r'업무 실적\(\d{4}\.\d{2}\.\d{2}\. ~ \d{4}\.\d{2}\.\d{2}\.\)',
                    f'업무 실적({(wed - timedelta(days=7)).strftime("%Y.%m.%d.")} ~ {(wed - timedelta(days=1)).strftime("%Y.%m.%d.")})', xml)
                xml = re.sub(
                    r'업무 계획\(\d{4}\.\d{2}\.\d{2}\. ~ \d{4}\.\d{2}\.\d{2}\.\)',
                    f'업무 계획({wed.strftime("%Y.%m.%d.")} ~ {(wed + timedelta(days=6)).strftime("%Y.%m.%d.")})', xml)
                # 김건양 연구실적 한 셀만 (04.24 구조: col 4, row 14)
                xml = replace_cell(xml, 4, 14, "테스트 연구실적 한 줄")
                files['Contents/section0.xml'] = xml.encode('utf-8')

                buf = _io.BytesIO()
                with zipfile.ZipFile(buf, 'w') as zout:
                    for name in order:
                        orig = infos[name]
                        zi = zipfile.ZipInfo(name, date_time=orig.date_time)
                        zi.compress_type = orig.compress_type
                        zi.external_attr = orig.external_attr
                        zi.create_system = orig.create_system
                        zi.create_version = orig.create_version
                        zi.extract_version = orig.extract_version
                        zi.flag_bits = orig.flag_bits
                        zi.extra = orig.extra
                        zout.writestr(zi, files[name])
                result = _patch_zip_flag_bits(buf.getvalue(), infos)
                st.download_button(
                    "💾 Lv3 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv3_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv3 = Lv2 + replace_cell 한 번 호출 (김건양 연구실적 r14 c4).")

            # Level 4: 셀 10개 수정 (색상 없음)
            if st.button("🔬 Lv4: 셀 10개 수정 (색상 X)",
                         use_container_width=True):
                import re
                from hwpx_exporter import _patch_zip_flag_bits, replace_cell
                src = debug_tpl.read_bytes()
                with zipfile.ZipFile(_io.BytesIO(src), 'r') as zin:
                    order = [i.filename for i in zin.infolist()]
                    infos = {i.filename: i for i in zin.infolist()}
                    files = {n: zin.read(n) for n in order}
                xml = files['Contents/section0.xml'].decode('utf-8')
                # 팀원 10명 연구실적 셀만 수정 (또는 업무실적)
                test_cells = [
                    (4, 2), (4, 5), (4, 8), (4, 11),   # 현장실증팀 연구실적
                    (4, 14), (4, 17), (4, 20), (4, 23), # 로봇기술팀 연구실적
                    (4, 25), (4, 26),                   # 최혜민/정지수 업무실적
                ]
                for i, (c, r) in enumerate(test_cells):
                    xml = replace_cell(xml, c, r, f"테스트 {i+1} 한 줄짜리")
                files['Contents/section0.xml'] = xml.encode('utf-8')

                buf = _io.BytesIO()
                with zipfile.ZipFile(buf, 'w') as zout:
                    for name in order:
                        orig = infos[name]
                        zi = zipfile.ZipInfo(name, date_time=orig.date_time)
                        zi.compress_type = orig.compress_type
                        zi.external_attr = orig.external_attr
                        zi.create_system = orig.create_system
                        zi.create_version = orig.create_version
                        zi.extract_version = orig.extract_version
                        zi.flag_bits = orig.flag_bits
                        zi.extra = orig.extra
                        zout.writestr(zi, files[name])
                result = _patch_zip_flag_bits(buf.getvalue(), infos)
                st.download_button(
                    "💾 Lv4 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv4_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv4 = 10개 셀 연달아 수정, 색상 override 없음.")

            # Level 5: Lv4 + 획득데이터에 파란색 색상 override
            if st.button("🔬 Lv5: Lv4 + 획득데이터 파란색",
                         use_container_width=True):
                import re
                from hwpx_exporter import (_patch_zip_flag_bits, replace_cell,
                                           ensure_blue_charpr)
                src = debug_tpl.read_bytes()
                with zipfile.ZipFile(_io.BytesIO(src), 'r') as zin:
                    order = [i.filename for i in zin.infolist()]
                    infos = {i.filename: i for i in zin.infolist()}
                    files = {n: zin.read(n) for n in order}
                xml = files['Contents/section0.xml'].decode('utf-8')
                header = files['Contents/header.xml'].decode('utf-8')
                header, blue_id = ensure_blue_charpr(header)
                # 10개 셀 + 4개 획득데이터 셀(파란색)
                test_cells = [
                    (4, 2), (4, 5), (4, 8), (4, 11),
                    (4, 14), (4, 17), (4, 20), (4, 23),
                    (4, 25), (4, 26),
                ]
                for i, (c, r) in enumerate(test_cells):
                    xml = replace_cell(xml, c, r, f"테스트 {i+1} 한 줄짜리")
                # 획득데이터 4개 파란색
                for i, (c, r) in enumerate([(4, 1), (4, 4), (4, 7), (4, 10)]):
                    xml = replace_cell(xml, c, r, f"획득 {i+1}",
                                       override_color_id=blue_id)
                files['Contents/section0.xml'] = xml.encode('utf-8')
                files['Contents/header.xml'] = header.encode('utf-8')

                buf = _io.BytesIO()
                with zipfile.ZipFile(buf, 'w') as zout:
                    for name in order:
                        orig = infos[name]
                        zi = zipfile.ZipInfo(name, date_time=orig.date_time)
                        zi.compress_type = orig.compress_type
                        zi.external_attr = orig.external_attr
                        zi.create_system = orig.create_system
                        zi.create_version = orig.create_version
                        zi.extract_version = orig.extract_version
                        zi.flag_bits = orig.flag_bits
                        zi.extra = orig.extra
                        zout.writestr(zi, files[name])
                result = _patch_zip_flag_bits(buf.getvalue(), infos)
                st.download_button(
                    "💾 Lv5 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv5_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv5 = Lv4 + 획득데이터 셀 4개에 파란색 override 추가.")

            # Level 6: 실제 build_report()를 빈 데이터로 호출
            if st.button("🔬 Lv6: build_report() 빈 데이터로 호출",
                         use_container_width=True):
                src = debug_tpl.read_bytes()
                result = build_report(
                    src, {},  # 빈 submissions
                    title_date=wed.strftime("%y.%m.%d."),
                    period_start=(wed - timedelta(days=7)).strftime("%Y.%m.%d."),
                    period_end=(wed - timedelta(days=1)).strftime("%Y.%m.%d."),
                    plan_start=wed.strftime("%Y.%m.%d."),
                    plan_end=(wed + timedelta(days=6)).strftime("%Y.%m.%d."),
                )
                st.download_button(
                    "💾 Lv6 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv6_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv6 = 실제 build_report() 함수를 빈 제출 데이터로 호출. 35+개 셀 모두 수정됨 (빈 값으로).")

            # Level 7: Lv6 + 멀티라인 텍스트
            if st.button("🔬 Lv7: build_report() 멀티라인 테스트 데이터",
                         use_container_width=True):
                src = debug_tpl.read_bytes()
                test_subs = {m['name']: {
                    'acquired_data': f'{m["name"]} 획득',
                    'research_done': f'1. {m["name"]} 연구 1\n2. 연구 2\n3. 연구 3',
                    'research_plan': f'{m["name"]} 계획\n- 세부 1\n- 세부 2',
                    'task_done': f'1. {m["name"]} 업무\n2. 업무 추가',
                    'task_plan': f'{m["name"]} 계획',
                    'smart_care_space_done': '스페이스 실적\n- A\n- B',
                    'smart_care_space_plan': '스페이스 계획',
                    'research_meeting': '회의 1\n회의 2',
                    'director_meeting': '주간 회의',
                    'mohw_weekly': '보산진 일정',
                } for m in TEAM_MEMBERS}
                result = build_report(
                    src, test_subs,
                    title_date=wed.strftime("%y.%m.%d."),
                    period_start=(wed - timedelta(days=7)).strftime("%Y.%m.%d."),
                    period_end=(wed - timedelta(days=1)).strftime("%Y.%m.%d."),
                    plan_start=wed.strftime("%Y.%m.%d."),
                    plan_end=(wed + timedelta(days=6)).strftime("%Y.%m.%d."),
                )
                st.download_button(
                    "💾 Lv7 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv7_{debug_tpl.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info("Lv7 = build_report() + 10명 모두에게 멀티라인 테스트 데이터.")

            # Level 8: 실제 데이터 + flag_bits 바이너리 패치 제거
            if st.button("🔬 Lv8: 실제 데이터, flag_bits 패치 없이",
                         use_container_width=True):
                import hwpx_exporter as he
                original_patch = he._patch_zip_flag_bits
                # 패치 함수를 no-op 으로 임시 교체
                he._patch_zip_flag_bits = lambda zip_bytes, original_infos: zip_bytes
                try:
                    subs_for = load_week(week)
                    last_week_str = (wed - timedelta(days=7)).strftime("%Y-%m-%d")
                    last_week_subs = load_week(last_week_str)
                    for name in MEMBER_NAMES:
                        cur = subs_for.get(name, {})
                        last = last_week_subs.get(name, {})
                        merged = dict(last)
                        for k, v in cur.items():
                            if v:
                                merged[k] = v
                        if merged:
                            subs_for[name] = merged

                    src = debug_tpl.read_bytes()
                    result = build_report(
                        src, subs_for,
                        title_date=wed.strftime("%y.%m.%d."),
                        period_start=(wed - timedelta(days=7)).strftime("%Y.%m.%d."),
                        period_end=(wed - timedelta(days=1)).strftime("%Y.%m.%d."),
                        plan_start=wed.strftime("%Y.%m.%d."),
                        plan_end=(wed + timedelta(days=6)).strftime("%Y.%m.%d."),
                    )
                    st.download_button(
                        "💾 Lv8 다운로드",
                        data=result,
                        file_name=f"DEBUG_Lv8_{debug_tpl.name}",
                        mime="application/octet-stream",
                        use_container_width=True,
                    )
                    st.info("Lv8 = production 과 동일하지만 _patch_zip_flag_bits 무효화.")
                finally:
                    he._patch_zip_flag_bits = original_patch

            st.markdown("---")
            st.caption("**팀원별 이분법**: 특정 팀원 데이터가 깨뜨리는지 확인")

            # 한 팀원씩 포함하여 테스트 (나머지는 지난주 fallback)
            test_members = st.multiselect(
                "실제 데이터 넣을 팀원 선택 (나머지는 빈 값)",
                MEMBER_NAMES,
                default=[],
                key="debug_members",
            )
            if st.button("🔬 Lv9: 선택한 팀원만 실제 데이터",
                         use_container_width=True):
                real_data = load_week(week)
                last_week_str = (wed - timedelta(days=7)).strftime("%Y-%m-%d")
                last_week_real = load_week(last_week_str)
                # merge 지난주 → 이번주
                full_data = {}
                for name in MEMBER_NAMES:
                    cur = real_data.get(name, {})
                    last = last_week_real.get(name, {})
                    merged = dict(last)
                    for k, v in cur.items():
                        if v:
                            merged[k] = v
                    if merged:
                        full_data[name] = merged

                # test_members 에 포함된 팀원만 실제 데이터 유지, 나머지는 빈 값
                filtered = {}
                for name in test_members:
                    if name in full_data:
                        filtered[name] = full_data[name]

                src = debug_tpl.read_bytes()
                result = build_report(
                    src, filtered,
                    title_date=wed.strftime("%y.%m.%d."),
                    period_start=(wed - timedelta(days=7)).strftime("%Y.%m.%d."),
                    period_end=(wed - timedelta(days=1)).strftime("%Y.%m.%d."),
                    plan_start=wed.strftime("%Y.%m.%d."),
                    plan_end=(wed + timedelta(days=6)).strftime("%Y.%m.%d."),
                )
                name_label = "_".join(test_members[:3]) if test_members else "empty"
                st.download_button(
                    "💾 Lv9 다운로드",
                    data=result,
                    file_name=f"DEBUG_Lv9_{name_label}.hwpx",
                    mime="application/octet-stream",
                    use_container_width=True,
                )
                st.info(f"Lv9 = {len(test_members)}명만 실제 데이터, 나머지 빈 값.")

    # 수요일 기준(보고일): 실적=지난주 수요일~이번주 화요일, 계획=이번주 수요일~다음주 화요일
    period_start = (wed - timedelta(days=7)).strftime("%Y.%m.%d.")  # 지난주 수요일
    period_end = (wed - timedelta(days=1)).strftime("%Y.%m.%d.")    # 이번주 화요일
    plan_start = wed.strftime("%Y.%m.%d.")                          # 이번주 수요일
    plan_end = (wed + timedelta(days=6)).strftime("%Y.%m.%d.")      # 다음주 화요일
    title_date = wed.strftime("%y.%m.%d.")

    c1, c2 = st.columns(2)
    with c1:
        period_start = st.text_input("실적 시작", period_start)
        period_end = st.text_input("실적 종료", period_end)
    with c2:
        plan_start = st.text_input("계획 시작", plan_start)
        plan_end = st.text_input("계획 종료", plan_end)

    title_date = st.text_input("제목 날짜", title_date)

    # 레포 루트 (streamlit_app.py의 부모의 부모)에서 HWPX 템플릿 찾기
    repo_root = Path(__file__).resolve().parent.parent
    template_files = sorted(repo_root.glob("돌봄로봇_업무보고*.hwpx"))
    template_path = st.selectbox(
        "템플릿 HWPX 파일",
        template_files,
        format_func=lambda p: p.name,
        index=len(template_files) - 1 if template_files else 0,
    ) if template_files else None

    uploaded = st.file_uploader("또는 템플릿 직접 업로드", type=["hwpx"])

    if st.button("📥 HWPX 생성 및 다운로드", type="primary", use_container_width=True):
        try:
            if uploaded is not None:
                template_bytes = uploaded.getvalue()
            elif template_path is not None:
                template_bytes = template_path.read_bytes()
            else:
                st.error("템플릿 HWPX를 선택하거나 업로드해주세요.")
                return

            submissions = load_week(week)

            # 미제출자는 지난주 내용 fallback (완전 미제출인 사람만)
            last_week_str = (wed - timedelta(days=7)).strftime("%Y-%m-%d")
            last_week_subs = load_week(last_week_str)
            fallback_used = []
            for name in MEMBER_NAMES:
                if name not in submissions and name in last_week_subs:
                    submissions[name] = last_week_subs[name]
                    fallback_used.append(name)
            if fallback_used:
                st.info(f"🔄 이번주 미제출 {len(fallback_used)}명은 지난주 내용으로 대체: "
                        f"{', '.join(fallback_used)}")

            result = build_report(
                template_bytes, submissions,
                title_date=title_date,
                period_start=period_start, period_end=period_end,
                plan_start=plan_start, plan_end=plan_end,
            )
            filename = f"돌봄로봇_업무보고({wed.strftime('%m.%d')})_취합본.hwpx"
            st.download_button(
                "💾 HWPX 다운로드",
                data=result,
                file_name=filename,
                mime="application/octet-stream",
                use_container_width=True,
            )
            st.success("생성 완료. 위 버튼으로 다운로드하세요.")
        except Exception as e:
            st.error(f"생성 실패: {e}")


def main():
    if not auth_gate():
        return

    with st.sidebar:
        st.caption(f"접속 모드: {'관리자' if st.session_state.get('is_admin') else '팀원'}")
        mode_options = ["업무보고 작성"]
        if st.session_state.get("is_admin"):
            mode_options.append("담당자 대시보드")
        mode = st.radio("메뉴", mode_options)
        st.divider()
        if st.button("로그아웃"):
            for k in ["authed", "is_admin"]:
                st.session_state.pop(k, None)
            st.query_params.clear()
            st.rerun()

    if mode == "업무보고 작성":
        member_page()
    else:
        admin_page()


if __name__ == "__main__":
    main()

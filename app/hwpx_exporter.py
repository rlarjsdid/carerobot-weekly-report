"""10명 업무보고 데이터를 HWPX 템플릿에 일괄 삽입하는 모듈.

- 기본 본문 색상: charPr 15 (검정, 템플릿에 이미 존재)
- 파란색: 런타임에 charPr 15를 복제 + textColor를 #0000FF로 수정해서
          header.xml에 동적 추가한 뒤 그 id 사용
- 셀 위치가 여러 테이블에 중복으로 존재할 수 있음 → nth 인덱스 지원
"""
import zipfile
import re
import io
import html
from team_config import TEAM_MEMBERS

CHARPR_BLACK = "15"
COLOR_HEX = {"black": "#000000", "blue": "#0000FF"}


def _sanitize_for_hwpx(text: str) -> str:
    """한글 HWPX 파서가 싫어하는 문자 정리.
    - 백슬래시(\\): 한글이 이스케이프 문자로 잘못 해석하는 것으로 추정. 제거.
    - 그 외 비정상적 whitespace: 정상 공백으로 치환.
    """
    if not text:
        return ""
    # 백슬래시 제거
    text = text.replace("\\", "")
    return text


def make_paragraph_xml(text: str, char_pr_id: str = CHARPR_BLACK,
                       is_first: bool = True) -> str:
    """원본 update_hwpx.py 와 동일한 하드코딩 방식.
    단순하고 검증된 구조라 한글이 안전하게 파싱."""
    escaped = html.escape(_sanitize_for_hwpx(text))
    pid = "2147483648" if is_first else "0"
    return (
        f'<hp:p id="{pid}" paraPrIDRef="27" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{char_pr_id}"><hp:t>{escaped}</hp:t></hp:run>'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1100" '
        f'textheight="1100" baseline="935" spacing="164" horzpos="0" '
        f'horzsize="31508" flags="393216"/>'
        f'</hp:linesegarray></hp:p>'
    )


def make_cell_content(text: str, char_pr_id: str = CHARPR_BLACK) -> str:
    """여러 줄 텍스트를 여러 <hp:p> 문단으로 변환."""
    lines = (text or "").splitlines() or [""]
    return "".join(
        make_paragraph_xml(line, char_pr_id=char_pr_id, is_first=(i == 0))
        for i, line in enumerate(lines)
    )


def find_cell_sublist(xml, col, row, nth=0):
    """nth번째로 나타나는 cellAddr col=col row=row 셀의 subList 내부 영역 반환."""
    addr_str = f'cellAddr colAddr="{col}" rowAddr="{row}"'
    pos = 0
    for _ in range(nth + 1):
        pos = xml.find(addr_str, pos)
        if pos == -1:
            return None, None
        addr_pos = pos
        pos += len(addr_str)
    tc_start = xml.rfind('<hp:tc ', 0, addr_pos)
    if tc_start == -1:
        return None, None
    sublist_start = xml.find('<hp:subList', tc_start)
    if sublist_start == -1 or sublist_start > addr_pos:
        return None, None
    sublist_content_start = xml.find('>', sublist_start) + 1
    sublist_end = xml.find('</hp:subList>', sublist_start)
    if sublist_end == -1:
        return None, None
    return sublist_content_start, sublist_end


def replace_cell(xml, col, row, text, override_color_id=None, nth=0):
    """셀 내용을 새 <hp:p> 블록으로 교체. 하드코딩된 구조 사용."""
    start, end = find_cell_sublist(xml, col, row, nth=nth)
    if start is None:
        return xml
    char_pr = override_color_id if override_color_id is not None else CHARPR_BLACK
    new_content = make_cell_content(text, char_pr_id=char_pr)
    return xml[:start] + new_content + xml[end:]


def ensure_blue_charpr(header_xml: str) -> tuple[str, str]:
    """header.xml에 파란색 charPr가 없으면 추가하고 해당 id 반환.
    이미 textColor=#0000FF인 charPr가 있으면 그걸 재사용."""
    existing = re.search(r'<hh:charPr\s+id="(\d+)"[^>]*textColor="#0000FF"', header_xml)
    if existing:
        return header_xml, existing.group(1)

    black_m = re.search(
        rf'<hh:charPr\s+id="{CHARPR_BLACK}"[^>]*?>.*?</hh:charPr>',
        header_xml, re.DOTALL)
    if not black_m:
        raise RuntimeError("템플릿에 기본 검정 charPr(15)가 없습니다.")
    black_xml = black_m.group(0)

    max_id = max(
        int(x) for x in re.findall(r'<hh:charPr\s+id="(\d+)"', header_xml)
    )
    new_id = str(max_id + 1)

    blue_xml = re.sub(
        r'id="\d+"', f'id="{new_id}"', black_xml, count=1
    )
    blue_xml = re.sub(
        r'textColor="#[0-9A-Fa-f]{6}"',
        'textColor="#0000FF"',
        blue_xml, count=1,
    )

    new_header = header_xml.replace(black_xml, black_xml + blue_xml)

    new_header = re.sub(
        r'(<hh:charProperties[^>]*itemCnt=")(\d+)(")',
        lambda m: f'{m.group(1)}{int(m.group(2)) + 1}{m.group(3)}',
        new_header, count=1,
    )
    return new_header, new_id


def build_report(template_bytes: bytes, submissions: dict,
                 title_date: str,
                 period_start: str, period_end: str,
                 plan_start: str, plan_end: str) -> bytes:
    """submissions = {이름: {필드키: 텍스트, ...}}"""
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        xml = zin.read('Contents/section0.xml').decode('utf-8')
        header = zin.read('Contents/header.xml').decode('utf-8')
        # 원본 ZipInfo 전체를 보존 (external_attr, create_system, create_version,
        # extract_version, flag_bits, date_time 등 한글이 검사할 가능성 있는 모든 메타)
        entry_order = [info.filename for info in zin.infolist()]
        original_infos = {info.filename: info for info in zin.infolist()}
        all_files = {name: zin.read(name) for name in zin.namelist()}

    header, blue_id = ensure_blue_charpr(header)
    color_to_id = {"black": CHARPR_BLACK, "blue": blue_id}

    xml = re.sub(
        r'과업별 업무 보고 \(\d{2}\.\d{2}\.\d{2}\.\)',
        f'과업별 업무 보고 ({title_date})',
        xml,
    )
    xml = re.sub(
        r'업무 실적\(\d{4}\.\d{2}\.\d{2}\. ~ \d{4}\.\d{2}\.\d{2}\.\)',
        f'업무 실적({period_start} ~ {period_end})',
        xml,
    )
    xml = re.sub(
        r'업무 계획\(\d{4}\.\d{2}\.\d{2}\. ~ \d{4}\.\d{2}\.\d{2}\.\)',
        f'업무 계획({plan_start} ~ {plan_end})',
        xml,
    )

    for m in TEAM_MEMBERS:
        data = submissions.get(m["name"], {})
        for field, spec in m["cells"].items():
            if spec is None:
                continue  # HWPX 매핑 보류 필드 (시트 저장만 됨)
            if field in ("research_done", "research_plan") and not m["has_research"]:
                continue
            if len(spec) == 2:
                col, row = spec
                color, nth = "black", 0
            elif len(spec) == 3:
                col, row, color = spec
                nth = 0
            elif len(spec) == 4:
                col, row, color, nth = spec
            else:
                raise ValueError(f"잘못된 셀 명세: {spec}")
            text = data.get(field, "")
            # 파란색으로 명시된 필드만 override. 검정은 원본 셀 charPr 유지.
            override = color_to_id["blue"] if color == "blue" else None
            xml = replace_cell(xml, col, row, text,
                               override_color_id=override,
                               nth=nth)

    all_files['Contents/section0.xml'] = xml.encode('utf-8')
    all_files['Contents/header.xml'] = header.encode('utf-8')

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zout:
        for name in entry_order:
            data = all_files[name]
            orig = original_infos[name]
            zinfo = zipfile.ZipInfo(name, date_time=orig.date_time)
            zinfo.compress_type = orig.compress_type
            zinfo.external_attr = orig.external_attr
            zinfo.create_system = orig.create_system
            zinfo.create_version = orig.create_version
            zinfo.extract_version = orig.extract_version
            zinfo.flag_bits = orig.flag_bits
            zinfo.extra = orig.extra
            zout.writestr(zinfo, data)

    # Python zipfile 이 writestr 과정에서 flag_bits를 자동 변경(특히 0x04 clear)
    # 한글은 일부 엔트리의 flag_bits=0x04 를 기대하므로 바이너리 레벨에서 복원
    raw = buf.getvalue()
    raw = _patch_zip_flag_bits(raw, original_infos)
    return raw


def _patch_zip_flag_bits(zip_bytes: bytes, original_infos: dict) -> bytes:
    """ZIP 로컬 파일 헤더와 중앙 디렉토리 엔트리의 flag_bits 필드를
    원본 ZipInfo.flag_bits 값으로 복원."""
    data = bytearray(zip_bytes)
    LFH_SIG = b'PK\x03\x04'
    CD_SIG = b'PK\x01\x02'

    # 로컬 파일 헤더 스캔
    pos = 0
    while True:
        idx = data.find(LFH_SIG, pos)
        if idx == -1:
            break
        # LFH 구조: sig(4) ver(2) flag(2) method(2) time(2) date(2) crc(4) csize(4) usize(4) nlen(2) elen(2) name extra
        name_len = int.from_bytes(data[idx + 26:idx + 28], 'little')
        extra_len = int.from_bytes(data[idx + 28:idx + 30], 'little')
        name = data[idx + 30:idx + 30 + name_len].decode('utf-8', errors='replace')
        if name in original_infos:
            target_flag = original_infos[name].flag_bits
            data[idx + 6:idx + 8] = target_flag.to_bytes(2, 'little')
        # 다음으로
        comp_size = int.from_bytes(data[idx + 18:idx + 22], 'little')
        pos = idx + 30 + name_len + extra_len + comp_size

    # 중앙 디렉토리 스캔
    pos = 0
    while True:
        idx = data.find(CD_SIG, pos)
        if idx == -1:
            break
        # CD 구조: sig(4) vermade(2) verneeded(2) flag(2) method(2) ...
        name_len = int.from_bytes(data[idx + 28:idx + 30], 'little')
        extra_len = int.from_bytes(data[idx + 30:idx + 32], 'little')
        cmt_len = int.from_bytes(data[idx + 32:idx + 34], 'little')
        name = data[idx + 46:idx + 46 + name_len].decode('utf-8', errors='replace')
        if name in original_infos:
            target_flag = original_infos[name].flag_bits
            data[idx + 8:idx + 10] = target_flag.to_bytes(2, 'little')
        pos = idx + 46 + name_len + extra_len + cmt_len

    return bytes(data)


def load_template(template_path: str) -> bytes:
    with open(template_path, 'rb') as f:
        return f.read()

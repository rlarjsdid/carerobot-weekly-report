# 트러블슈팅 기록

이 프로젝트를 만들면서 마주친 **비자명한 오류와 해결 방법**을 남겨둠.
다음에 비슷한 문제 만나면 여기부터 찾아볼 것.

---

## 1. HWPX 생성본이 한글에서 "파일 손상" 에러로 안 열림

가장 오래 물고 늘어진 문제. 원인이 **5개**가 쌓여 있었음.

### 증상
관리자가 HWPX 다운로드 → 한글에서 열려고 하면 "파일이 손상되었거나 올바른
형식이 아닙니다" 류의 에러. 워낭오피스·한컴오피스 모두 동일.

### 근본 원인들 (발견된 순서)

#### 1-1. `charPrIDRef` 하드코딩 오류
- `hwpx_exporter.py`의 `make_paragraph_xml()`에서 `charPrIDRef="45"`로
  하드코딩되어 있었음 (옛날 `update_hwpx.py` 에서 복사해온 값)
- id=45는 원본 템플릿의 `header.xml`에 정의되어 있지 않은 ID → 파서 실패
- **수정**: id=15 (템플릿 기본 검정)로 변경. 필요시 `ensure_blue_charpr()`
  이 파란색(id=28) charPr를 header에 주입하고 ID 반환

#### 1-2. `mimetype` 압축 문제 (OPC 스펙 위반)
- HWPX는 OPC(Open Packaging Convention) 포맷. 그 스펙상 `mimetype` 파일은:
  - ZIP 의 **첫 엔트리**
  - **STORED(무압축)** 방식
  여야 함
- Python `zipfile.writestr()` 기본은 DEFLATE 압축 → 한글이 거부
- **수정**: 원본 ZIP 각 엔트리의 `compress_type` 을 그대로 복사해서 재포장

#### 1-3. ZIP `flag_bits` 메타데이터 덮어쓰기
- Python `zipfile`은 엔트리 쓸 때 `flag_bits`(Local File Header / Central
  Directory 에 저장되는 플래그)를 자체 기준으로 재계산함
- 한글은 이 플래그를 엄격하게 검증 → 원본과 다르면 거부
- **수정**: `_patch_zip_flag_bits()` — Python이 패키징한 뒤 **바이트 단위**로
  ZIP을 뒤져서 원본 `flag_bits` 복원
  - Local File Header: offset +6
  - Central Directory: offset +8

#### 1-4. 🐛 탭 문자(`\t`)가 특정 팀원 데이터에 섞여 있었음
- **증상**: "10명 중 특정 한 명 데이터만 포함하면 파일이 안 열림"
- **범인**: 한벼리의 연구실적 텍스트에 **탭 문자 4개** 포함 — 다른 문서에서
  번호 매겨가며 복붙하는 과정에서 들어감. 눈에는 안 보이지만 `repr()`
  찍으면 `\t` 로 나옴
- HWPX XML의 `<hp:t>` 태그 안에 raw 탭이 들어가면 한글 파서가 깨짐
- **수정**: `_sanitize_for_hwpx()` 추가 — 탭 → 스페이스 4개 변환, 백슬래시
  제거, 제어문자 필터 (단 `\n`·`\r` 은 유지)
- 커밋: `5f48a63 한벼리 HWPX 불가 범인: 탭(\t) 문자 4개`

#### 1-5. 자간(letter spacing)이 셀마다 안 맞음
- **증상**: 파일은 열리지만 특정 셀의 글자 간격이 비정상적으로 벌어지거나 좁음
- **원인**: `make_paragraph_xml()` 이 `horzsize="31508"` 같은 값을
  하드코딩 — 셀마다 폭이 다른데 이 값은 고정
- **수정**: `extract_cell_lineseg()` 추가 — 각 셀의 원본 `<hp:linesegarray>`
  를 뽑아서 그 셀 내부 문단에 그대로 재사용
- 커밋: `4ff67bb HWPX 자간 문제: 각 셀 원본의 linesegarray 보존`

### 디버깅 접근법: 단계별 이분법 (Lv1 ~ Lv9)

원인을 좁히기 위해 관리자 대시보드에 **디버그 버튼 9개**를 임시로 만들어서,
수정 단계를 한 번에 하나씩 추가함:

| Lv | 내용 | 목적 |
|----|------|------|
| 1 | ZIP 재포장만 (내용 수정 0) | ZIP 재포장 자체가 문제인가? |
| 2 | 제목/기간 날짜만 regex 치환 | 간단한 텍스트 치환도 깨뜨리나? |
| 3 | 셀 1개만 수정 | `replace_cell()` 한 번이 문제인가? |
| 4 | 셀 10개 수정 (색상 없음) | 연속 수정이 누적되어 깨지나? |
| 5 | Lv4 + 획득데이터 파란색 | charPr override 가 문제인가? |
| 6 | 실제 `build_report()` + 빈 데이터 | 35+개 셀 전체 수정이 문제? |
| 7 | Lv6 + 멀티라인 테스트 데이터 | 멀티라인 텍스트가 문제? |
| 8 | 실제 데이터, flag_bits 패치 OFF | flag_bits 패치가 오히려 문제? |
| **9** | **팀원별로 골라서 실제 데이터** | **특정 팀원 데이터가 문제?** |

결정적 발견은 **Lv9**. "한벼리를 빼면 열리고, 포함시키면 안 열린다"
→ 데이터 자체가 범인이란 걸 확신 → 한벼리 텍스트 `repr()` 찍어보니 탭 발견.

이 디버그 도구는 원인 규명 후 제거됨 (커밋 `13da847`).
**다시 필요하면**: `git show 13da847~1:app/streamlit_app.py` 로 복원 가능.

### 재발 시 체크리스트

1. 한글에서 "파일 손상" 에러 뜨면:
   - [ ] `_sanitize_for_hwpx()` 가 최근 제출 데이터에 호출되는지 확인
   - [ ] 구글시트에서 각 팀원 데이터 복사 → Python `repr()` 로 출력
         → `\t`, `\u00xx` 등 비정상 문자 육안 확인
   - [ ] 이분법: `app/streamlit_app.py` 에 한 팀원씩만 포함하는 임시 버튼
         만들어서 어느 팀원이 범인인지 특정
2. 자간만 이상하면:
   - [ ] `extract_cell_lineseg()` 가 셀별 원본 lineseg 받고 있는지 확인
   - [ ] 하드코딩된 `horzsize` 값이 남아있지 않은지 grep
3. 색상이 이상하면:
   - [ ] `charPrIDRef` 값이 실제 `header.xml` 에 있는 id 인지 확인
   - [ ] `ensure_blue_charpr()` 가 `header.xml` 수정본을 반환하는지 확인

---

## 2. 획득 데이터 필드 prefix 자동 추가

### 증상
기술팀 팀원이 획득 데이터 칸에 그냥 "없음" 만 입력 → 보고서에 덩그러니
"없음" 만 찍힘 → 보기 이상함

### 해결
`build_report()` 에서 `acquired_data` 필드 특수 처리:
- 빈 값 → `"획득 데이터:"` (라벨만)
- `"획득 데이터"` 로 시작하지 않는 텍스트 → `"획득 데이터: {입력}"` 자동 prefix
- 이미 `"획득 데이터"` 로 시작하면 그대로 둠 (중복 방지)

커밋: `ddfcd44`

---

## 3. Word(.docx) 백업 출력 포기

### 경위
HWPX 생성이 계속 실패하던 초기에는 Word 를 "한컴·워드 모두 열리는 백업
포맷"으로 만들어 둠. 하지만:

1. **양식 문제**: 셀 병합, 세로쓰기 라벨("사업단공통확인사항" 같은 것),
   폭 비율이 원본 HWPX와 똑같이 안 나옴
2. **사용자 피드백**: "첫 페이지부터 글자가 세로로 써져있고 양식도 안맞아"
3. HWPX 원인 다 잡고 나니 백업이 불필요해짐

### 최종 결정
Word/PDF 출력 전체 제거 (커밋 `13da847`) — HWPX 만 지원. `reportlab`,
`python-docx` 의존성도 `requirements.txt` 에서 뺌.

---

## 4. HWPX 내부 구조 메모 (참고용)

다시 HWPX 만질 일 있으면 이 섹션부터 읽기.

### 파일 구조
HWPX = ZIP + XML. `unzip` 이나 `zipfile` 로 열어볼 수 있음:

```
├── mimetype                     ← STORED 필수, 첫 엔트리
├── META-INF/
│   └── manifest.xml
├── Contents/
│   ├── header.xml               ← 스타일/색상(charPr) 정의
│   └── section0.xml             ← 본문 테이블, 문단
├── settings.xml
└── ...
```

### 주요 XML 규칙

#### `charPrIDRef` (문자 속성 = 색상·폰트·크기)
- `header.xml` 의 `<hh:charProperties>` 에 정의된 id 참조
- 우리 템플릿 기준:
  - `id=15` : 검정 기본
  - `id=28` : 파란색 (획득 데이터용)
- `<hp:run charPrIDRef="15">` 이런 식으로 문단 안에 지정

#### paragraph id
- 각 셀의 **첫 문단**: `id="2147483648"` (= 2^31)
- 나머지 문단: `id="0"`
- 틀리면 한글이 파일을 거부할 수 있음

#### `<hp:linesegarray>` (자간·줄높이)
- 문단 끝에 항상 붙는 필수 요소
- `horzsize` 는 셀 폭에 종속 → **절대 하드코딩 금지**
- `extract_cell_lineseg()` 로 셀마다 원본 값 뽑아서 재사용해야 함

### 파서를 깨뜨리는 문자 (반드시 sanitize)
- `\t` (탭, U+0009) → 스페이스 4개로 치환
- `\` (백슬래시) → 제거 (escape 해석 위험)
- 제어 문자 U+0000 ~ U+001F (단 `\n`, `\r` 제외) → 제거
- 전부 `_sanitize_for_hwpx()` 에서 처리됨

### `ZipInfo.flag_bits` 핵심
```python
# 원본 flag_bits 를 살려서 쓰려면:
zi = zipfile.ZipInfo(name, date_time=orig.date_time)
zi.compress_type = orig.compress_type
zi.flag_bits = orig.flag_bits         # ← 이게 핵심
# ... 나머지 속성 복사 ...
zout.writestr(zi, files[name])

# 하지만 Python이 writestr 시점에 flag_bits 를 덮어씀 →
# 마지막에 바이너리 패치 필요:
result = _patch_zip_flag_bits(buf.getvalue(), original_infos)
```

---

## 관련 커밋 참조

| 커밋 | 내용 |
|------|------|
| `b3f7554` | 백슬래시 제거 (초기 가설, 부분적으로 맞음) |
| `5f48a63` | 탭 문자 제거 — **결정적 수정** |
| `ddfcd44` | 획득 데이터 prefix 자동 추가 |
| `4ff67bb` | 자간(linesegarray) 셀별 원본 보존 |
| `13da847` | PDF/Word 및 디버그 도구 제거 (정리 커밋) |

---

## 핵심 교훈 (2026년 4월 김건양 메모)

1. **"되는 파일 vs 안 되는 파일"의 바이트 diff 부터 비교**. XML 파싱보다
   빠르고 정확함.
2. **원인이 하나가 아닐 수 있다**. HWPX 안 열림 문제는 5가지 원인이
   순서대로 나타났음. 하나 고치고 "여전히 안 되네" 에 포기하지 말 것.
3. **특정 팀원 데이터만 의심스러울 때는 이분법이 최강**. 전체 파싱 로직
   분석보다 빨랐음. Lv9 버튼 만들어서 한벼리 범인이란 거 5분 만에 찾음.
4. **`repr()` 은 거짓말 안 함**. 눈에 안 보이는 탭·제어문자 찾을 때 필수.
5. **백업 포맷(Word/PDF) 유지보수 비용을 과소평가하지 말 것**. 양식
   동일성 맞추기가 생각보다 어려워서 결국 다 버림.

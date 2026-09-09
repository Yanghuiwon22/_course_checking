
from pathlib import Path

# ── 기준 경로 (실행 위치 무관) ────────────────────────────────────
BASE_DIR    = Path(__file__).parent          # 과제점검/
PROJECT_DIR = BASE_DIR.parent               # 프원실/

# ── 파일/폴더 경로 ────────────────────────────────────────────────
# 현재 과목 명단. 과목을 바꾸면 이 줄만 교체한다.
# 프원실로 되돌릴 때: BASE_DIR / "프원실(2026)_과제확인.xlsx"
ROSTER_PATH = PROJECT_DIR / "input" / "스마트팜데이터과학_참여자 목록.xlsx"
CREDS_PATH  = BASE_DIR / "credentials.json"  # 구글시트 인증키
LMS_DIR     = PROJECT_DIR / "lms 제출물"
GITHUB_DIR  = PROJECT_DIR / "github 과제점검"
OUTPUT_DIR  = BASE_DIR / "output"
LMS_ZIP_DIR = PROJECT_DIR / "input"      # LMS에서 받은 제출물 ZIP을 넣는 곳

# 제출물 재편에 쓸 명단. 보통 ROSTER_PATH 와 같다.
SUBMISSION_ROSTER = ROSTER_PATH

# ── 수강생 명단 열 번호 (프원실(2026)_과제확인.xlsx) ─────────────────────────────────
COL_STUDENT_ID   = 4   # D열: 학번
COL_STUDENT_NAME = 5   # E열: 이름
COL_GITHUB_ID    = 6   # F열: GitHub ID
COL_CLONE_CMD    = 13  # M열: clone 명령어

# ── Google Sheets ─────────────────────────────────────────────────
SPREADSHEET_ID = '1xiR5yvUDt_vceA6Jr4a-p2sGTYhIjmdOSKNwC3xFJ5M'
SHEET_GID      = 0


# -- 이번 주 점검 설정 ----
TARGET_COURSE = '스마트팜데이터과학'
TARGET_WEEK   = '2주차'
HW_NUM        = 2

# LMS만 하는지, Github도 하는지 확인(과목에 따라 다름)
CHECK_LMS     = True
CHECK_GITHUB  = False

# ── 과제별 내용 점검 기준 (강의자료/과제XX.txt 기반) ─────────────────
# 새 과제를 점검할 때 여기에 항목을 추가한다.
from datetime import datetime as _dt
RUBRICS = {
    1: {
        "kind": "text",
        "min_chars": 1500,
        "filename_re": r"데싸_과제01_(?P<name>[^_]+)_(?P<sid>\d+)_자기소개",
        "deadline": _dt(2026, 9, 6, 12, 0, 0),
        # 5개 필수 항목 완전성은 키워드 자동판정을 시도했으나 오탐률이 높아(표본 4/4 오탐)
        # 제외함. 형식(글자수/파일명/기한)만 자동 판정하고, 내용 충실도는 사람이 직접 확인한다.
    },
    2: {
        "kind": "code",
        "ext": [".ipynb"],
        "deadline": _dt(2026, 9, 6, 12, 0, 0),
        # 과제02.txt: "함수 작성, 함수 호출 등을 구현" — 정의/호출 여부만 확인
    },
}

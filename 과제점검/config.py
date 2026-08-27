"""
config.py
─────────
경로·상수 중앙 관리 모듈.
모든 스크립트가 이 파일에서 경로와 상수를 import한다.
"""

from pathlib import Path

# ── 기준 경로 (실행 위치 무관) ────────────────────────────────────
BASE_DIR    = Path(__file__).parent          # 과제점검/
PROJECT_DIR = BASE_DIR.parent               # 프원실/

# ── 파일/폴더 경로 ────────────────────────────────────────────────
ROSTER_PATH = BASE_DIR / "프원실(2026)_과제확인.xlsx"
CREDS_PATH  = BASE_DIR / "credentials.json"
LMS_DIR     = PROJECT_DIR / "lms 제출물"
GITHUB_DIR  = PROJECT_DIR / "github 과제점검"
OUTPUT_DIR  = BASE_DIR / "output"

# ── 수강생 명단 열 번호 (1-based) ─────────────────────────────────
COL_STUDENT_ID   = 4   # D열: 학번
COL_STUDENT_NAME = 5   # E열: 이름
COL_GITHUB_ID    = 6   # F열: GitHub ID
COL_CLONE_CMD    = 13  # M열: clone 명령어

# ── Google Sheets ─────────────────────────────────────────────────
SPREADSHEET_ID = '1wGzelPw95eADLmglfaOtA3ZumQDM8yu43lwEO4AtPiU'
SHEET_GID      = 865348684

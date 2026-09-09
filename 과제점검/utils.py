"""
utils.py
────────
공통 유틸리티 모듈.
여러 스크립트에서 공유하는 함수와 상수를 모아 놓는다.
"""

import re


def normalize_name(raw: str) -> str:
    """엑셀 학생 이름을 Windows 파일명으로 정규화."""
    s = re.sub(r"\s+", " ", str(raw)).strip()
    s = re.sub(r'[<>:"/\\|?*]', "_", s)
    return s.rstrip(" .") or "이름없음"


# ── Excel 색상 팔레트 ─────────────────────────────────────────────
# build_report_XX.py 에서도 동일하게 사용한다.
COLORS = {
    "dark":     "1A1A2E",   # 제목 배경
    "navy":     "2E4057",   # 헤더 배경
    "green":    "D4EDDA",   # 정상/제출
    "yellow":   "FFF3CD",   # 경고/일부제출
    "orange":   "FFE5CC",   # 실행오류
    "pink":     "F8D7DA",   # 오류/이슈
    "gray":     "E2E3E5",   # 미제출
    "white":    "FFFFFF",
    "stripe":   "EAF4F4",   # 짝수행 배경
    "issue":    "721C24",   # 이슈 글자색
    "sum":      "048A81",   # 요약 배경
    "lavender": "D6C9F0",   # 코드체크 그룹 헤더 (build_report 전용)
}


# ── 수강생 명단 ───────────────────────────────────────────────────
# LMS '참여자 목록'처럼 헤더가 있는 파일은 자동 인식하고,
# 헤더가 없는 파일(프원실 과제확인표)은 열 번호를 받아 처리한다.
HEADER_NAME = ("이름", "성명")
HEADER_ID   = ("아이디", "학번", "id")


def load_roster_rows(path, col_name: int | None = None,
                     col_id: int | None = None) -> list[tuple[str, str]]:
    """명단에서 [(이름, 학번), ...] 를 순서대로 읽는다. 교수·조교 행은 제외."""
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []

    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    i_name = next((i for i, h in enumerate(header) if h in HEADER_NAME), None)
    i_id   = next((i for i, h in enumerate(header) if h in HEADER_ID), None)
    i_role = next((i for i, h in enumerate(header) if h == "역할"), None)

    if i_name is None or i_id is None:          # 헤더 없음 → 넘겨받은 열 번호 사용
        if col_name is None or col_id is None:
            raise ValueError(f"{path}: 헤더를 찾지 못했고 열 번호도 지정되지 않았습니다.")
        i_name, i_id, i_role = col_name - 1, col_id - 1, None

    out: list[tuple[str, str]] = []
    for row in rows[1:]:
        if i_name >= len(row) or not row[i_name]:
            continue
        if i_role is not None and i_role < len(row) and row[i_role]:
            if str(row[i_role]).strip() != "학생":
                continue
        name = normalize_name(row[i_name]).split(" ")[0]
        sid  = str(row[i_id]).strip() if i_id < len(row) and row[i_id] else ""
        out.append((name, sid))
    return out


def student_dirname(name: str, sid: str) -> str:
    """학생 폴더명 규칙. 동명이인 구분을 위해 항상 이름(학번) 형태를 쓴다."""
    return f"{name}({sid})" if sid else name


def resolve_student_dir(root, name: str, sid: str):
    """실제로 존재하는 학생 폴더를 찾는다. 이름(학번) → 이름 순으로 시도."""
    from pathlib import Path as _P
    root = _P(root)
    for cand in (student_dirname(name, sid), name):
        if (root / cand).is_dir():
            return root / cand
    return root / student_dirname(name, sid)     # 없으면 규칙상 경로를 반환

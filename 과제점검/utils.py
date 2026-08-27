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

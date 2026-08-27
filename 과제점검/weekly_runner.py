"""
weekly_runner.py
────────────────
매주 일요일 자정, Cowork 예약 작업에 의해 자동 실행됩니다.

1. 학기 시작일(SEMESTER_START) 기준으로 현재 주차를 계산합니다.
2. check_main.py 의 TARGET_WEEK 값을 해당 주차로 수정합니다.
3. check_main.py 를 실행하고 stdout / stderr 를 캡처합니다.
4. 결과(성공/실패, 단계별 오류)를 출력합니다.
"""

import re
import subprocess
import sys
from datetime import date
from pathlib import Path

# ── 설정 ──────────────────────────────────────────────
SEMESTER_START = date(2026, 3, 2)   # 1주차 시작일 (월요일)
CHECK_MAIN     = Path(__file__).parent / "check_main.py"
TIMEOUT_SEC    = 600                # 10분 제한
# ──────────────────────────────────────────────────────


def calc_week() -> int:
    """오늘 날짜 기준 학기 주차를 반환한다."""
    delta = (date.today() - SEMESTER_START).days
    week  = delta // 7 + 1
    return max(1, week)


def update_target_week(week: int) -> str:
    """check_main.py 의 TARGET_WEEK 를 수정하고, 이전 값을 반환한다."""
    text     = CHECK_MAIN.read_text(encoding="utf-8")
    pattern  = r"(TARGET_WEEK\s*=\s*')[^']*(')"
    new_val  = f"{week}주차"
    new_text = re.sub(pattern, rf"\g<1>{new_val}\g<2>", text)

    old_match = re.search(pattern, text)
    old_val   = old_match.group(0) if old_match else "(알 수 없음)"

    CHECK_MAIN.write_text(new_text, encoding="utf-8")
    return old_val


def run_check_main() -> tuple[int, str, str]:
    """check_main.py 를 실행하고 (returncode, stdout, stderr) 를 반환한다."""
    result = subprocess.run(
        [sys.executable, str(CHECK_MAIN)],
        cwd=CHECK_MAIN.parent,
        capture_output=True,
        text=True,
        timeout=TIMEOUT_SEC,
        encoding="utf-8",
        errors="replace",
    )
    return result.returncode, result.stdout, result.stderr


def summarize_issues(stdout: str, stderr: str) -> list[str]:
    """stdout / stderr 에서 오류 줄을 추출해 반환한다."""
    issues = []
    error_patterns = re.compile(
        r"(❌|Error|Exception|Traceback|FAIL|실패|오류)",
        re.IGNORECASE,
    )
    for line in (stdout + "\n" + stderr).splitlines():
        if error_patterns.search(line):
            issues.append(line.strip())
    return issues


def main():
    week = calc_week()
    print(f"\n{'='*52}")
    print(f"  weekly_runner  |  실행 날짜: {date.today()}  |  {week}주차")
    print(f"{'='*52}")

    # ── TARGET_WEEK 수정 ──────────────────────────────
    old_val = update_target_week(week)
    print(f"\n[1] TARGET_WEEK 업데이트: {old_val!r} → '{week}주차'")

    # ── check_main.py 실행 ────────────────────────────
    print(f"[2] check_main.py 실행 중... (최대 {TIMEOUT_SEC}초)")
    try:
        rc, stdout, stderr = run_check_main()
    except subprocess.TimeoutExpired:
        print(f"\n⏰ 타임아웃: {TIMEOUT_SEC}초 초과")
        sys.exit(1)

    print(stdout)
    if stderr:
        print("--- STDERR ---", file=sys.stderr)
        print(stderr, file=sys.stderr)

    # ── 이슈 요약 ─────────────────────────────────────
    issues = summarize_issues(stdout, stderr)
    print(f"\n{'─'*52}")
    if issues:
        print(f"⚠  발견된 이슈 ({len(issues)}건):")
        for i, line in enumerate(issues, 1):
            print(f"  {i}. {line}")
    else:
        print("✅ 이슈 없음. 모든 단계 정상 완료.")
    print(f"{'─'*52}")

    sys.exit(rc)


if __name__ == "__main__":
    main()

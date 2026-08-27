"""
check_main.py
─────────────
프원실 과제 점검 전체 자동화 진입점.

실행 전 아래 두 변수를 수정하세요:
    TARGET_COURSE : LMS 과목명 (예: '프로그래밍원리와실습')
    TARGET_WEEK   : 주차 문자열 (예: '7주차')  ← hw 번호 자동 파싱

실행 순서:
    1. LMS 자동 로그인 & 제출물 다운로드   (auto_LMS_Task.auto_login)
    2. GitHub 저장소 clone / pull          (clone_from_excel)
    3. LMS 제출 여부 확인 → Google Sheets  (check_submissions)
    4. GitHub 과제 코드 점검 → xlsx 리포트 (auto_check_github)
    5. 최종 통합 리포트 생성               (build_report + data/hw{XX}.py — 존재할 때만)
"""

import re
import sys
from pathlib import Path

import check_submissions
import auto_login
import clone_from_excel
import auto_check_github
import build_report
from config import BASE_DIR


# ────────────────────────────────────────────────
# ✏️  여기만 수정하세요
TARGET_COURSE = '프로그래밍원리와실습'
TARGET_WEEK   = '17주차'          # 예: '7주차', '8주차', '10주차' ...

def try_build_report(hw_num: int):
    """data/hw{hw2}.py 가 있으면 build_report.run() 실행, 없으면 안내만."""
    hw2       = f"{hw_num:02d}"
    data_path = BASE_DIR / "data" / f"hw{hw2}.py"

    if not data_path.exists():
        print(f"\n⚠  [5단계 건너뜀] data/hw{hw2}.py 없음.")
        print(f"   data/hw_template.py 복사 후 학생 데이터를 채우세요.")
        return

    print(f"\n📋 [5단계] 최종 통합 리포트 생성 (data/hw{hw2}.py)")
    try:
        build_report.run(hw_num)
        print("✅ 통합 리포트 생성 완료")
    except Exception as e:
        print(f"❌ 통합 리포트 생성 실패: {e}", file=sys.stderr)


def main():

    print("=" * 50)
    print(f"  프원실 과제 점검 자동화 시작")
    print(f"  과목: {TARGET_COURSE} / 과제: ({TARGET_WEEK})주차")
    print("=" * 50)

    ROSTER_PATH = Path('..') / '과제점검' / '프원실(2026)_과제확인.xlsx'
    STUDENT_DIR = Path(r"E:\2026_1학기_프원실\기말_코드_모음")


    auto_login.init_student_dirs(ROSTER_PATH, STUDENT_DIR, 5)
    #
    # # ── 1단계: LMS 자동 로그인 & 제출물 다운로드 ────────────
    # print(f"\n🌐 [1단계] LMS 자동 로그인 & 제출물 다운로드")
    # try:
    #     auto_login.main(TARGET_COURSE, TARGET_WEEK)
    #     print("✅ LMS 다운로드 완료")
    # except Exception as e:
    #     print(f"❌ LMS 단계 실패: {e}", file=sys.stderr)
    #
    # # ── 2단계: GitHub 저장소 clone / pull ───────────────────
    # print(f"\n🔗 [2단계] GitHub 저장소 clone / pull")
    # try:
    #     clone_from_excel.main()
    #     print("✅ clone / pull 완료")
    # except Exception as e:
    #     print(f"❌ clone 단계 실패: {e}", file=sys.stderr)

    # # ── 3단계: LMS 제출 여부 확인 ───────────────────────────
    # print(f"\n📊 [3단계] LMS 제출 여부 확인 → Google Sheets 업데이트")
    # try:
    #     check_submissions.main()
    #     print("✅ LMS 제출 확인 완료")
    # except Exception as e:
    #     print(f"❌ 제출 확인 단계 실패: {e}", file=sys.stderr)

    # # ── 4단계: GitHub 과제 코드 점검 ────────────────────────
    # print(f"\n🔍 [4단계] GitHub 과제{hw2} 코드 점검")
    # try:
    #     hw_num = 11
    #     out_path = auto_check_github.run(hw_num)
    #     print(f"✅ GitHub 점검 리포트 저장: {out_path}")
    # except Exception as e:
    #     print(f"❌ GitHub 점검 단계 실패: {e}", file=sys.stderr)

    # # ── 5단계: 최종 통합 리포트 (존재할 때만) ───────────────
    # try_build_report(hw_num)
    #
    # print("\n" + "=" * 50)
    # print("  전체 점검 완료!")
    # print("=" * 50)


if __name__ == '__main__':
    main()

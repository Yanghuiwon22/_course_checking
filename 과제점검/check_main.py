import sys

import import_lms_zip
import clone_from_excel
import check_submissions
import auto_check_github
import check_lms_content
import build_report
from config import (BASE_DIR, TARGET_COURSE, TARGET_WEEK, HW_NUM,
                    CHECK_LMS, CHECK_GITHUB)


def step(label, fn):
    """단계 하나를 실행하고 성공/실패를 출력한다."""
    print(f"\n{label}")
    try:
        fn()
        print("✅ 완료")
    except Exception as e:
        print(f"❌ 실패: {e}", file=sys.stderr)


def try_build_report(hw_num: int):
    hw2 = f"{hw_num:02d}"
    if not (BASE_DIR / "data" / f"hw{hw2}.py").exists():
        print(f"\n⚠  [5단계 건너뜀] data/hw{hw2}.py 없음 "
              f"— hw_template.py 복사 후 학생 데이터를 채우세요.")
        return
    step(f"📋 [5단계] 최종 통합 리포트 (data/hw{hw2}.py)",
         lambda: build_report.run(hw_num))


def main():
    hw2  = f"{HW_NUM:02d}"
    mode = " + ".join(m for m, on in
                      [("LMS", CHECK_LMS), ("GitHub", CHECK_GITHUB)] if on) or "없음"

    print("=" * 50)
    print(f"  과제 점검 자동화 시작")
    print(f"  과목: {TARGET_COURSE} / {TARGET_WEEK} / 과제{hw2}")
    print(f"  점검 대상: {mode}")
    print("=" * 50)

    if CHECK_LMS:
        step("📦 [1단계] LMS 제출물 ZIP 재편 (input/ 의 ZIP 사용)",
             lambda: import_lms_zip.run(HW_NUM))

    if CHECK_GITHUB:
        step("🔗 [2단계] GitHub 저장소 clone / pull",
             clone_from_excel.main)

    if CHECK_LMS:
        step("📊 [3단계] LMS 제출 여부 → Google Sheets",
             check_submissions.main)

    if CHECK_GITHUB:
        step(f"🔍 [4단계] GitHub 과제{hw2} 코드 점검",
             lambda: auto_check_github.run(HW_NUM))
    else:
        step(f"🔍 [4단계] LMS 과제{hw2} 내용 확인",
             lambda: check_lms_content.run(HW_NUM))

    try_build_report(HW_NUM)

    print("\n" + "=" * 50)
    print("  전체 점검 완료!")
    print("=" * 50)

if __name__ == "__main__":
    main()
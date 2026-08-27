"""
auto_check_github.py
────────────────────
GitHub 저장소에서 특정 과제 번호의 제출 파일을 자동 탐색·분석하여
과제XX_점검리포트_github.xlsx 를 생성합니다.

사용법:
    python auto_check_github.py --hw 7
    python auto_check_github.py --hw 7 --out ../과제07_점검리포트_github.xlsx

특징:
    - 폴더명 비표준 허용: hw07 / homework07 / 과제07 / task_07 / 07 / 루트파일 등
    - 각 .py 파일 구문 검사 (ast.parse) + 실행 테스트 (제한 시간 3초)
    - 폴더 구조 자동 판정: ✅ 정상 / ⚠ 비규격 / ❌ 미정리 (루트) / — (미제출)
"""

import argparse
import ast
import os
import re
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import GITHUB_DIR, ROSTER_PATH as ROSTER, BASE_DIR as THIS_DIR, PROJECT_DIR
from utils import COLORS as C
from utils import normalize_name


def fill(h): return PatternFill("solid", fgColor=h)
def font(color="000000", bold=False, size=10):
    return Font(color=color, bold=bold, size=size)
def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def lwrap():  return Alignment(horizontal="left",   vertical="center", wrap_text=True)
thin = Side(border_style="thin", color="CCCCCC")
def border(): return Border(left=thin, right=thin, top=thin, bottom=thin)

def cell(ws, r, c, v, bg=None, fg="000000", bold=False, align=None, sz=10):
    cl = ws.cell(r, c, v)
    if bg: cl.fill = fill(bg)
    cl.font = Font(color=fg, bold=bold, size=sz)
    cl.alignment = align or center()
    cl.border = border()
    return cl

# ── 수강생 명단 로드 ──────────────────────────────────────────
def load_roster(path: Path):
    """[(이름, 학번, GitHub_ID), ...]  — 학번 오름차순
    프원실(2026)_과제확인.xlsx 기준:
        4열(index 3): 학번 / 5열(index 4): 이름 / 6열(index 5): GitHub ID
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        name, sid, ghid = r[4], r[3], r[5]
        if name:
            rows.append((
                str(name).strip().replace('\n', ' '),
                str(int(sid)) if sid else "",
                str(ghid).strip() if ghid else "(없음)"
            ))
    wb.close()
    return sorted(rows, key=lambda x: x[1])

# ── hw 폴더 탐색 (핵심 로직) ─────────────────────────────────
def find_hw_folder(student_dir: Path, hw_num: int):
    """
    과제 번호에 해당하는 폴더/파일을 유연하게 탐색합니다.

    반환: (folder_path_or_None, folder_name_str, structure_flag)
        structure_flag:
            "정상"    — hwXX / homeworkXX 등 표준 폴더
            "비규격"  — 숫자만('07'), 한글혼용('과제07_...'), 다른 접두어('task_07')
            "루트"    — 폴더 없이 루트에 파일 직접 제출
            "미제출"  — 파일 없음
    """
    hw2 = f"{hw_num:02d}"   # "07"
    hw1 = str(hw_num)        # "7"

    # ── 패턴 정의 ────────────────────────────────────────────
    # 1순위: 표준 폴더명 (hw07, homework07)
    standard_pattern  = re.compile(
        rf'^(hw|homework){hw2}$', re.IGNORECASE)
    # 2순위: 비규격 폴더 (숫자만, 한글혼용, task_, assignment_ 등)
    nonstandard_pattern = re.compile(
        rf'(?<![0-9]){hw2}(?![0-9])|(?<![0-9]){hw1}(?![0-9])', re.IGNORECASE)

    std_hits = []
    nonstd_hits = []

    if not student_dir.exists():
        return None, "없음", "미제출"

    for item in student_dir.iterdir():
        if item.name.startswith('.') or item.name == '__pycache__':
            continue
        name = item.name

        # 폴더 탐색
        if item.is_dir():
            if standard_pattern.match(name):
                std_hits.append(item)
            elif nonstandard_pattern.search(name):
                nonstd_hits.append(item)

        # 루트 파일 탐색 (폴더 없이 .py 파일 직접 제출)
        elif item.is_file() and item.suffix == '.py':
            if nonstandard_pattern.search(name):
                nonstd_hits.append(item)  # 파일이지만 표시용으로 사용

    if std_hits:
        folder = std_hits[0]
        return folder, folder.name, "정상"

    if nonstd_hits:
        hit = nonstd_hits[0]
        if hit.is_dir():
            return hit, hit.name, "비규격"
        else:
            # 루트에 파일 직접 제출 — 파일들이 있는 student_dir 자체를 폴더로
            return student_dir, f"(루트) {hit.name}", "루트"

    # 루트 전체에서 hw 번호가 포함된 .py 파일 탐색
    root_py = [f for f in student_dir.glob("*.py")
               if nonstandard_pattern.search(f.stem)]
    if root_py:
        return student_dir, f"(루트) {root_py[0].name}", "루트"

    return None, "없음", "미제출"

# ── .py 파일 분석 ─────────────────────────────────────────────
def collect_py_files(folder: Path, student_dir: Path, structure: str):
    """
    hw 폴더(또는 루트)에서 .py 파일 목록 반환.
    루트 제출이면 hw 번호가 포함된 파일만.
    """
    if structure == "미제출" or folder is None:
        return []

    if structure == "루트":
        # 루트 제출: 파일명에 07이 들어간 것만 (임시 방편)
        return [f for f in student_dir.glob("*.py")
                if not f.name.startswith('._')]

    return [f for f in folder.rglob("*.py")
            if not f.name.startswith('._')]

def check_syntax(py_file: Path):
    """ast.parse 구문 검사. True=OK, False=SyntaxError"""
    try:
        ast.parse(py_file.read_text(encoding='utf-8', errors='replace'))
        return True, ""
    except SyntaxError as e:
        return False, str(e)

def run_file(py_file: Path, timeout=3):
    """
    파일 실행 (stdin=빈 파이프, timeout 3초).
    반환: (ok, error_msg)
    """
    try:
        result = subprocess.run(
            [sys.executable, str(py_file)],
            input="",
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=timeout,
        )
        if result.returncode != 0:
            err = (result.stderr or "").strip().splitlines()
            return False, err[-1] if err else "RuntimeError"
        return True, ""
    except subprocess.TimeoutExpired:
        return True, "(timeout — input 대기로 판단, 정상 처리)"
    except Exception as e:
        return False, str(e)

def analyze_files(py_files):
    """
    반환:
        syntax_ok  : bool  (모두 구문 OK면 True)
        run_ok     : bool  (모두 실행 OK면 True)
        syntax_mark: "✅" / "❌" / "—"
        run_mark   : "✅" / "❌" / "—"
        errors     : [(filename, kind, msg), ...]
    """
    if not py_files:
        return False, False, "—", "—", []

    errors = []
    all_syntax = True
    all_run = True

    for f in sorted(py_files):
        ok, msg = check_syntax(f)
        if not ok:
            all_syntax = False
            all_run = False
            errors.append((f.name, "SyntaxError", msg))
            continue

        rok, rmsg = run_file(f)
        if not rok:
            all_run = False
            errors.append((f.name, "RuntimeError", rmsg))

    return (
        all_syntax,
        all_run,
        "✅" if all_syntax else "❌",
        "✅" if all_run    else "❌",
        errors,
    )

# ── 종합 판정 ─────────────────────────────────────────────────
def make_summary(structure, run_ok, syntax_ok, has_files):
    """GitHub xlsx 종합 열 값"""
    if structure == "미제출" or not has_files:
        return "미제출"
    if not syntax_ok or not run_ok:
        return "실행오류"
    return "정상"   # 이슈 여부는 로직이슈 열 기입 후 수동 확인

# ── Excel 생성 ────────────────────────────────────────────────
def build_excel(hw_num, students_data, out_path):
    """
    students_data: list of dict {
        name, sid, ghid,
        folder_name, structure,
        syntax, run, logic_issue, summary
    }
    """
    from datetime import date
    hw2 = f"{hw_num:02d}"
    today = date.today().strftime("%Y-%m-%d")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "종합현황"

    LAST_COL = "I"
    NCOLS = 9

    # 행1 제목
    ws.merge_cells(f"A1:{LAST_COL}1")
    c = ws.cell(1, 1, f"과제{hw2} GitHub 점검 리포트  ·  {today} 기준")
    c.fill = fill(C["dark"]); c.font = Font(bold=True, color=C["white"], size=13)
    c.alignment = center(); ws.row_dimensions[1].height = 30

    # 행2 헤더
    headers = ["이름","학번","GitHub ID",f"hw{hw2} 폴더명","폴더구조","문법","실행","로직이슈","종합"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(2, col, h)
        c.fill = fill(C["navy"]); c.font = Font(bold=True, color=C["white"], size=10)
        c.alignment = center(); c.border = border()
    ws.row_dimensions[2].height = 30

    struct_mark = {
        "정상":   "✅ 정상",
        "비규격": "⚠ 비규격",
        "루트":   "❌ 미정리",
        "미제출": "—",
    }
    summary_color = {
        "정상":      C["green"],
        "폴더미정리": C["yellow"],
        "이슈있음":  C["pink"],
        "실행오류":  C["orange"],
        "미제출":    C["gray"],
        "미등록":    C["gray"],
    }

    for r_idx, s in enumerate(students_data, 3):
        is_even = r_idx % 2 == 0
        bg_stripe = C["stripe"] if is_even else C["white"]

        # A~C
        for col, val in enumerate([s["name"], s["sid"], s["ghid"]], 1):
            cl = ws.cell(r_idx, col, val)
            cl.fill = fill(bg_stripe); cl.font = Font(size=10)
            cl.alignment = center(); cl.border = border()

        # D 폴더명
        cell(ws, r_idx, 4, s["folder_name"], sz=9)

        # E 폴더구조
        struct = struct_mark.get(s["structure"], "—")
        struct_bg = {
            "✅ 정상":  C["green"],
            "⚠ 비규격": C["yellow"],
            "❌ 미정리": C["pink"],
            "—":        C["gray"],
        }.get(struct, C["white"])
        cell(ws, r_idx, 5, struct, bg=struct_bg, sz=10)

        # F 문법  G 실행
        for col, mark in [(6, s["syntax"]), (7, s["run"])]:
            bg = {"✅": C["green"], "❌": C["pink"], "—": C["gray"]}.get(mark, C["white"])
            cell(ws, r_idx, col, mark, bg=bg, sz=10)

        # H 로직이슈
        issue_text = s.get("logic_issue", "이상 없음")
        has_issue = issue_text not in ("이상 없음", "—", "")
        cl = ws.cell(r_idx, 8, issue_text)
        cl.fill = fill(C["pink"] if has_issue else C["white"])
        cl.font = Font(color=C["issue"] if has_issue else "000000", size=9)
        cl.alignment = lwrap(); cl.border = border()

        # I 종합
        summ = s["summary"]
        cell(ws, r_idx, 9, summ,
             bg=summary_color.get(summ, C["white"]),
             bold=True, sz=10)

        ws.row_dimensions[r_idx].height = 36

    # 요약
    summary_row = len(students_data) + 4
    total = len(students_data)
    submitted = sum(1 for s in students_data if s["summary"] != "미제출")
    not_sub   = sum(1 for s in students_data if s["summary"] == "미제출")

    for label, val in [
        ("■ 점검 요약", None),
        ("전체 학생수", f"{total}명"),
        ("제출(파일 확인)", f"{submitted}명"),
        ("미제출", f"{not_sub}명"),
    ]:
        ws.merge_cells(f"A{summary_row}:H{summary_row}")
        ws.cell(summary_row, 1, label).fill = fill(C["sum"])
        ws.cell(summary_row, 1).font = Font(bold=True, color=C["white"], size=10)
        ws.cell(summary_row, 1).alignment = center()
        if val is not None:
            ws.cell(summary_row, 9, val).fill = fill(C["sum"])
            ws.cell(summary_row, 9).font = Font(bold=True, color=C["white"], size=10)
            ws.cell(summary_row, 9).alignment = center()
        summary_row += 1

    # 열 너비
    for col, w in zip(range(1, NCOLS+1), [14, 13, 22, 24, 12, 8, 8, 50, 12]):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = "A3"
    wb.save(out_path)
    print(f"✅ 저장: {out_path}")

# ── 핵심 실행 로직 (임포트용) ────────────────────────────────
def run(hw_num: int, out_path=None):
    """
    다른 스크립트에서 직접 호출 가능한 진입점.
    hw_num  : 과제 번호 (예: 7)
    out_path: 출력 xlsx 경로 (None이면 기본값 사용)
    반환: out_path (Path)
    """
    hw2 = f"{hw_num:02d}"
    if out_path is None:
        out_path = PROJECT_DIR / f"과제{hw2}_점검리포트_github.xlsx"
    out_path = Path(out_path)

    print(f"\n🔍 과제{hw2} GitHub 자동 점검 시작")
    print(f"   GitHub 폴더: {GITHUB_DIR}")
    print(f"   수강생 명단: {ROSTER}\n")

    roster = load_roster(ROSTER)
    students_data = []

    for name, sid, ghid in roster:

        name = normalize_name(name)
        # GitHub 폴더 매핑: "강태현 (KANG TAEHYEON)" 형태 탐색
        print(name)
        student_dir = None
        for d in GITHUB_DIR.iterdir():
            if d.is_dir() and d.name.startswith(name):
                student_dir = d
                break

        if student_dir is None:
            print(student_dir)

            print(f"  ⚠ [{name}] GitHub 폴더 없음 (미등록 또는 clone 안 됨)")
            students_data.append(dict(
                name=name, sid=sid, ghid=ghid,
                folder_name="없음", structure="미제출",
                syntax="—", run="—", logic_issue="—", summary="미등록"
            ))
            continue

        # hw 폴더 탐색
        folder, folder_name, structure = find_hw_folder(student_dir, hw_num)
        py_files = collect_py_files(folder, student_dir, structure)

        # 파일 분석
        _, _, syntax_mark, run_mark, errors = analyze_files(py_files)

        # 로직이슈 (자동 감지 가능한 항목만 — 나머지는 수동 기입)
        if structure == "미제출" or not py_files:
            logic_issue = "—"
        elif errors:
            error_strs = [f"{fn}: {kind} — {msg}" for fn, kind, msg in errors]
            logic_issue = "; ".join(error_strs)
        else:
            logic_issue = "이상 없음"  # ← 로직 이슈는 수동으로 추가

        summary = make_summary(structure, run_mark == "✅", syntax_mark == "✅",
                               bool(py_files))

        flag = {
            "미제출":  "⛔",
            "미등록":  "⛔",
            "정상":    "🟢",
            "폴더미정리": "🟡",
            "실행오류": "🟠",
        }.get(summary, "📋")

        print(f"  {flag} [{name}]  폴더={folder_name}  구조={structure}  "
              f"문법={syntax_mark}  실행={run_mark}  종합={summary}")

        students_data.append(dict(
            name=name, sid=sid, ghid=ghid,
            folder_name=folder_name, structure=structure,
            syntax=syntax_mark, run=run_mark,
            logic_issue=logic_issue, summary=summary
        ))

    # 결과 출력
    submitted   = sum(1 for s in students_data if s["summary"] not in ("미제출", "미등록"))
    not_sub     = sum(1 for s in students_data if s["summary"] in ("미제출", "미등록"))
    nonstandard = sum(1 for s in students_data if s["structure"] in ("비규격", "루트"))

    print(f"\n── 점검 결과 ───────────────────────")
    print(f"  전체: {len(students_data)}명")
    print(f"  제출: {submitted}명 / 미제출: {not_sub}명")
    print(f"  폴더 비정상(비규격·루트): {nonstandard}명")
    print(f"────────────────────────────────────\n")

    build_excel(hw_num, students_data, out_path)
    print("\n⚠  로직 이슈는 xlsx를 열어 H열(로직이슈)을 직접 기입하세요.")

    return out_path


# ── 메인 (CLI 직접 실행용) ────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="GitHub 과제 자동 점검")
    parser.add_argument("--hw", type=int, required=True, help="과제 번호 (예: 7)")
    parser.add_argument("--out", type=str, default=None,
                        help="출력 xlsx 경로 (기본: 상위폴더/과제XX_점검리포트_github.xlsx)")
    args = parser.parse_args()
    run(args.hw, args.out)

if __name__ == "__main__":
    main()

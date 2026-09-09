"""
check_hw_rubric.py
────────────────────
과제별 실제 기준(config.RUBRICS)에 맞춰 LMS 제출 내용을 판정하고,
결과를 구글 시트에 과제 번호별 새 탭으로 정리한다.

check_lms_content.py 가 '파일을 열어 텍스트를 뽑아내는' 범용 도구라면,
이 스크립트는 그 텍스트를 실제 과제 요구사항과 대조해 정상/이슈를 가른다.

사용법:
    python check_hw_rubric.py --hw 1
"""
from __future__ import annotations

import argparse
import re
import zipfile
from datetime import datetime
from pathlib import Path

import gspread

from check_lms_content import analyze_student_files, extract_ipynb, EXTRACTORS
from config import (RUBRICS, ROSTER_PATH, LMS_DIR, LMS_ZIP_DIR, CREDS_PATH,
                    SPREADSHEET_ID, COL_STUDENT_NAME, COL_STUDENT_ID)
from utils import load_roster_rows, resolve_student_dir, COLORS as C

# 추출기가 "참고용"이라고 밝힌 값(예: PrvText 미리보기)은 글자수 판정에 신뢰할 수 없다.
SOFT_ISSUE_MARK = "참고용"


def find_hw_zip(zip_dir: Path, hw_num: int) -> Path | None:
    """input/ 에서 이 과제 번호가 들어간 ZIP을 찾는다 (지각 판정용 원본 타임스탬프)."""
    hw2 = f"{hw_num:02d}"
    cands = [z for z in zip_dir.glob("*.zip") if f"과제{hw2}" in z.name]
    return max(cands, key=lambda f: f.stat().st_mtime) if cands else None


def load_submit_times(zip_path: Path | None) -> dict[str, datetime]:
    """ZIP 안 파일들의 최종 수정시각 → {최상위 폴더명: 제출시각}."""
    if zip_path is None:
        return {}
    out: dict[str, datetime] = {}
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            if info.filename.endswith("/"):
                continue
            top = info.filename.split("/")[0]
            dt = datetime(*info.date_time)
            if top not in out or dt > out[top]:
                out[top] = dt
    return out


def find_code_file(hw_dir: Path, exts: list[str]) -> tuple[Path | None, str | None]:
    """지정 확장자 파일을 찾는다. 없으면 내용이 ipynb JSON으로 보이는 파일을 폴백으로 찾는다
    (예: Colab에서 받으면서 확장자가 .txt로 바뀐 경우)."""
    files = [f for f in hw_dir.iterdir() if f.is_file()]
    for f in files:
        if f.suffix.lower() in exts:
            return f, None
    for f in files:
        try:
            head = f.read_text(encoding="utf-8", errors="replace")[:80]
        except Exception:
            continue
        if head.lstrip().startswith("{") and '"nbformat"' in head:
            return f, f"확장자가 {'/'.join(exts)} 아님 ({f.name})"
    return None, None


def extract_full_code(path: Path) -> tuple[str, str]:
    """규칙 판정용 전체 코드 텍스트 (미리보기 300자 제한 없이)."""
    if path.suffix.lower() == ".ipynb" or path.name.lower().endswith(".ipynb.txt"):
        return extract_ipynb(path)
    try:
        return path.read_text(encoding="utf-8", errors="replace"), ""
    except Exception as e:
        return "", f"읽기 실패({type(e).__name__})"


def check_code_requirements(code: str) -> list[str]:
    """함수 정의(def)와 그 함수를 호출한 흔적이 있는지 간단히 확인한다."""
    issues = []
    defs = re.findall(r"^\s*def\s+(\w+)\s*\(", code, re.MULTILINE)
    if not defs:
        return ["함수 정의(def) 없음"]

    called = any(
        not code[code.rfind("\n", 0, m.start()) + 1: m.start()].lstrip().startswith("def")
        for name in set(defs)
        for m in re.finditer(rf"\b{re.escape(name)}\s*\(", code)
    )
    if not called:
        issues.append("정의한 함수를 호출한 흔적 없음")
    return issues


def extract_full_text(hw_dir: Path, files: list[dict]) -> str:
    """규칙 판정용 전체 본문 텍스트 (미리보기 300자 제한 없이 다시 읽는다)."""
    if not files:
        return ""
    path = hw_dir / files[0]["file"]
    extractor = EXTRACTORS.get(path.suffix.lower())
    if not extractor:
        return ""
    text, _ = extractor(path)
    return text


def check_sections(text: str, sections: list[tuple[str, list[str]]]) -> list[str]:
    """필수 항목별 키워드가 본문에 있는지 확인한다 (완전 자동 판정 아님 — 의심 신호)."""
    return [label for label, patterns in sections
           if not any(re.search(p, text) for p in patterns)]


def judge_student(name: str, sid: str, files: list[dict], hw_dir: Path,
                  rubric: dict, submit_time: datetime | None) -> tuple[str, list[str]]:
    """한 학생을 rubric 기준으로 판정한다. 반환: (종합판정, 이슈목록)"""
    if not files:
        return "미제출", []

    issues: list[str] = []
    kind = rubric.get("kind", "text")

    if kind == "code":
        target, ext_issue = find_code_file(hw_dir, rubric.get("ext", [".ipynb"]))
        if target is None:
            issues.append(f"필수 파일 없음 ({'/'.join(rubric.get('ext', ['.ipynb']))})")
        else:
            if ext_issue:
                issues.append(ext_issue)
            code, err = extract_full_code(target)
            if err:
                issues.append(f"{err} — 수동 확인 필요")
            else:
                issues.extend(check_code_requirements(code))
    else:
        hard_fail = [f for f in files
                    if f["issue"] and SOFT_ISSUE_MARK not in f["issue"]]
        if hard_fail:
            issues.append("본문 추출 실패 — 수동 확인 필요")

        total_chars = sum(f["chars"] for f in files)
        min_chars = rubric.get("min_chars")
        if min_chars and not hard_fail and total_chars < min_chars:
            issues.append(f"글자수 부족 ({total_chars}자 / 최소 {min_chars}자)")

        fname_re = rubric.get("filename_re")
        if fname_re and not re.search(fname_re, files[0]["file"]):
            issues.append(f"파일명 규칙 위반 ({files[0]['file']})")

        sections = rubric.get("sections")
        if sections and not hard_fail:
            missing = check_sections(extract_full_text(hw_dir, files), sections)
            if missing:
                issues.append(f"내용 누락 의심: {', '.join(missing)} — 실제 서술 여부 확인 필요")

    deadline = rubric.get("deadline")
    if deadline and submit_time and submit_time > deadline:
        issues.append(f"지각 ({submit_time:%Y-%m-%d %H:%M})")

    verdict = "이상없음" if not issues else "이슈"
    return verdict, issues


def run(hw_num: int) -> None:
    hw2, folder_name = f"{hw_num:02d}", f"hw{hw_num:02d}"
    rubric = RUBRICS.get(hw_num)
    if rubric is None:
        raise ValueError(f"config.RUBRICS 에 과제{hw2} 기준이 없습니다. "
                         f"강의자료/과제{hw2}.txt 를 보고 항목을 추가해 주세요.")

    roster = load_roster_rows(ROSTER_PATH, COL_STUDENT_NAME, COL_STUDENT_ID)
    submit_times = load_submit_times(find_hw_zip(LMS_ZIP_DIR, hw_num))

    kind_desc = (f"{rubric['min_chars']}자 이상" if rubric.get("kind", "text") == "text"
                else "함수 정의/호출 확인")
    print(f"🔍 과제{hw2} 내용 점검 (기준: {kind_desc}, 명단 {len(roster)}명)")

    rows, n_ok, n_issue, n_empty = [], 0, 0, 0
    for name, sid in roster:
        student_dir = resolve_student_dir(LMS_DIR, name, sid)
        hw_dir = student_dir / folder_name
        files = analyze_student_files(hw_dir)

        top = next((k for k in submit_times if k.startswith(name)), None)
        verdict, issues = judge_student(name, sid, files, hw_dir, rubric,
                                        submit_times.get(top))
        if verdict == "미제출":
            n_empty += 1
        elif verdict == "이슈":
            n_issue += 1
        else:
            n_ok += 1
        rows.append((f"{name}({sid})", verdict, "; ".join(issues)))

    print(f"✅ 이상없음 {n_ok} / 이슈 {n_issue} / 미제출 {n_empty}")

    write_sheet(hw2, rows)


def write_sheet(hw2: str, rows: list[tuple[str, str, str]]) -> None:
    """같은 스프레드시트에 '과제{NN}' 탭을 만들어(있으면 비우고) 결과를 쓴다."""
    gc = gspread.service_account(filename=str(CREDS_PATH))
    sh = gc.open_by_key(SPREADSHEET_ID)
    title = f"과제{hw2}"

    try:
        ws = sh.worksheet(title)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=len(rows) + 5, cols=4)

    header = ["학생", "종합판정", "이슈"]
    ws.update([header] + [list(r) for r in rows])

    def hex_to_rgb(h: str) -> dict:
        h = h.lstrip("#")
        return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}

    VERDICT_BG = {"이상없음": C["green"], "이슈": C["yellow"], "미제출": C["gray"]}
    fmt_requests = [{
        "range": "A1:C1",
        "format": {"backgroundColor": {"red": 0.18, "green": 0.25, "blue": 0.34},
                  "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}},
    }]
    for i, (_, verdict, _issue) in enumerate(rows, start=2):
        bg = VERDICT_BG.get(verdict)
        if bg:
            fmt_requests.append({"range": f"B{i}", "format": {"backgroundColor": hex_to_rgb(bg)}})
    ws.batch_format(fmt_requests)   # 한 번의 API 호출로 모든 서식 적용 (쓰기 할당량 절약)

    ws.freeze(rows=1)


def main():
    ap = argparse.ArgumentParser(description="과제 내용 규칙 점검 → 구글시트 새 탭")
    ap.add_argument("--hw", type=int, required=True, help="과제 번호 (예: 1)")
    args = ap.parse_args()
    run(args.hw)


if __name__ == "__main__":
    main()

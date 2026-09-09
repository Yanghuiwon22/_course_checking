"""
check_lms_content.py
─────────────────────
GitHub 점검을 하지 않는 과제(CHECK_GITHUB=False)를 위해 LMS 제출물의
실제 내용을 확인한다. auto_check_github.py는 '존재 여부'만 보던 3단계와
달리 파일을 열어 내용을 뽑아낸다.

- .py            : 구문 검사(ast.parse) + 3초 실행 테스트 (auto_check_github 재사용)
- .hwp           : OLE 'PrvText' 스트림에서 미리보기 텍스트 추출 (없으면 수동 확인 안내)
- .hwpx / .docx  : OOXML(zip+XML) 본문 텍스트 추출
- .txt / .html   : 텍스트로 읽기 (html은 태그 제거)
- 그 외          : 파일명 · 크기만 기록, 수동 확인 안내

결과: 과제{NN}_점검리포트_lms.xlsx
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from auto_check_github import check_syntax, run_file
from config import HW_NUM, LMS_DIR, PROJECT_DIR, ROSTER_PATH, COL_STUDENT_ID, COL_STUDENT_NAME
from utils import load_roster_rows, resolve_student_dir, COLORS as C

PREVIEW_LEN = 300   # 미리보기로 보여줄 글자 수


# ── 파일 형식별 텍스트 추출기 ────────────────────────────────────
def _hwp5txt_path() -> str | None:
    """pyhwp의 hwp5txt 실행 파일 경로. 없으면 None."""
    import shutil
    exe = Path(sys.executable).parent / "Scripts" / "hwp5txt.exe"
    if exe.exists():
        return str(exe)
    return shutil.which("hwp5txt")


def extract_hwp(path: Path) -> tuple[str, str]:
    """
    .hwp 전체 본문 텍스트 추출.

    자동 추출(pyhwp/PrvText)이 모두 실패한 파일은 사람이 원문을 직접 읽고
    옆에 "<원본파일명>.manual.txt" 로 저장해 둘 수 있다. 이 파일이 있으면
    항상 최우선으로 사용한다 (사람이 확인한 값이 가장 정확하므로).

    OLE 'PrvText' 스트림은 미리보기용이라 ~1000자 안팎에서 잘려 있어
    글자수 기준(예: 1500자 이상) 판정에 쓸 수 없다. pyhwp(hwp5txt)로 전체
    본문을 추출하고, pyhwp가 없거나 실패할 때만 PrvText로 대체한다.
    """
    manual = path.with_name(path.name + ".manual.txt")
    if manual.exists():
        text = manual.read_text(encoding="utf-8", errors="replace").strip()
        return text, ""

    hwp5txt = _hwp5txt_path()
    if hwp5txt:
        try:
            r = subprocess.run([hwp5txt, str(path)], capture_output=True,
                              text=True, encoding="utf-8", errors="replace", timeout=20)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip(), ""
        except Exception:
            pass   # 아래 PrvText 대체 경로로 진행

    try:
        import olefile
    except ImportError:
        return "", "olefile 미설치 — 수동 확인 필요"
    try:
        ole = olefile.OleFileIO(str(path))
        if not ole.exists("PrvText"):
            ole.close()
            return "", "본문 추출 실패, 미리보기도 없음 — 수동 확인 필요"
        text = ole.openstream("PrvText").read().decode("utf-16-le", errors="replace")
        ole.close()
        return text.strip(), "미리보기 텍스트만 추출됨(전체 본문 아님) — 글자수 참고용으로만 사용"
    except Exception as e:
        return "", f"읽기 실패({type(e).__name__}) — 수동 확인 필요"


def extract_ooxml(path: Path) -> tuple[str, str]:
    """.hwpx / .docx 등 zip+XML 포맷: 본문 텍스트 노드를 모아 반환한다."""
    try:
        with zipfile.ZipFile(path) as zf:
            chunks = []
            for name in zf.namelist():
                low = name.lower()
                if not low.endswith(".xml"):
                    continue
                if not any(k in low for k in ("section", "document", "body", "content")):
                    continue
                try:
                    root = ET.fromstring(zf.read(name))
                except ET.ParseError:
                    continue
                chunks.append(" ".join(t.strip() for t in root.itertext() if t.strip()))
            text = " ".join(chunks).strip()
            return text, ("" if text else "본문 텍스트를 찾지 못함 — 수동 확인 필요")
    except zipfile.BadZipFile:
        return "", "zip 형식이 아님 — 수동 확인 필요"
    except Exception as e:
        return "", f"읽기 실패({type(e).__name__}) — 수동 확인 필요"


def extract_html(path: Path) -> tuple[str, str]:
    raw = path.read_text(encoding="utf-8", errors="replace")
    text = re.sub(r"<[^>]+>", " ", raw)
    return re.sub(r"\s+", " ", text).strip(), ""


def extract_txt(path: Path) -> tuple[str, str]:
    return path.read_text(encoding="utf-8", errors="replace").strip(), ""


def extract_ipynb(path: Path) -> tuple[str, str]:
    """Jupyter/Colab 노트북: 모든 셀의 소스를 이어 붙인다 (코드+마크다운)."""
    import json
    try:
        nb = json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception as e:
        return "", f"노트북 파싱 실패({type(e).__name__}) — 수동 확인 필요"

    chunks = []
    for cell in nb.get("cells", []):
        src = cell.get("source", "")
        if isinstance(src, list):
            src = "".join(src)
        # Colab 기본 안내 셀 등에 이미지가 data URI로 통째로 박히는 경우가 있다
        # (예: ![logo.webp](data:image/webp;base64,...) — 수만~수십만 자까지 부풀림).
        # 글자수 판정을 왜곡하므로 제거한다.
        src = re.sub(r"data:[\w/+.-]+;base64,[A-Za-z0-9+/=]+", "[이미지 생략]", src)
        if src.strip():
            chunks.append(src)
    text = "\n\n".join(chunks).strip()
    return text, ("" if text else "빈 노트북 — 수동 확인 필요")


EXTRACTORS = {
    ".hwp": extract_hwp,
    ".hwpx": extract_ooxml,
    ".docx": extract_ooxml,
    ".html": extract_html,
    ".htm": extract_html,
    ".txt": extract_txt,
    ".ipynb": extract_ipynb,
}


# ── 학생 1명 분석 ────────────────────────────────────────────────
def analyze_student_files(hw_dir: Path) -> list[dict]:
    """hw 폴더 안 파일들을 검사해 결과 리스트로 반환한다."""
    if not hw_dir.is_dir():
        return []

    rows = []
    for f in sorted(hw_dir.iterdir()):
        if not f.is_file():
            continue
        if f.name.lower().endswith(".manual.txt"):
            continue   # 사람이 넣어둔 수동 추출본 — 원본 파일 처리 시 자동으로 함께 쓰임
        ext = f.suffix.lower()

        if ext == ".py":
            ok, err = check_syntax(f)
            if ok:
                ok, err = run_file(f)
            rows.append({
                "file": f.name, "kind": "코드",
                "preview": "", "issue": "" if ok else err,
                "chars": len(f.read_text(encoding="utf-8", errors="replace")),
            })
            continue

        extractor = EXTRACTORS.get(ext)
        if extractor:
            text, err = extractor(f)
            rows.append({"file": f.name, "kind": ext.lstrip("."),
                        "preview": text[:PREVIEW_LEN], "issue": err, "chars": len(text)})
        else:
            rows.append({"file": f.name, "kind": ext.lstrip(".") or "?",
                        "preview": "", "issue": "지원하지 않는 형식 — 수동 확인 필요",
                        "chars": f.stat().st_size})
    return rows


# ── 리포트 생성 ──────────────────────────────────────────────────
def build_excel(students_data: list[dict], out_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "종합현황"

    headers = ["이름", "학번", "폴더", "파일명", "종류", "글자수", "미리보기 / 확인 필요 사유"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=C["navy"])
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

    r = 2
    for s in students_data:
        if not s["files"]:
            ws.cell(r, 1, s["name"]); ws.cell(r, 2, s["sid"])
            ws.cell(r, 3, "미제출").fill = PatternFill("solid", fgColor=C["gray"])
            r += 1
            continue
        for f in s["files"]:
            ws.cell(r, 1, s["name"]); ws.cell(r, 2, s["sid"]); ws.cell(r, 3, s["folder"])
            ws.cell(r, 4, f["file"]); ws.cell(r, 5, f["kind"]); ws.cell(r, 6, f["chars"])
            cell = ws.cell(r, 7, f["issue"] or f["preview"])
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if f["issue"]:
                cell.fill = PatternFill("solid", fgColor=C["pink"])
            r += 1

    for i, w in enumerate((12, 12, 20, 34, 8, 8, 70), 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)


def run(hw_num: int = HW_NUM, out_path: Path | None = None) -> Path:
    hw2, folder_name = f"{hw_num:02d}", f"hw{hw_num:02d}"
    roster = load_roster_rows(ROSTER_PATH, COL_STUDENT_NAME, COL_STUDENT_ID)

    print(f"🔍 LMS 제출 내용 확인: {folder_name} ({len(roster)}명)")

    students_data, n_issue, n_empty = [], 0, 0
    for name, sid in roster:
        student_dir = resolve_student_dir(LMS_DIR, name, sid)
        files = analyze_student_files(student_dir / folder_name)
        if not files:
            n_empty += 1
        elif any(f["issue"] for f in files):
            n_issue += 1
        students_data.append({"name": name, "sid": sid,
                              "folder": student_dir.name, "files": files})

    out_path = Path(out_path) if out_path else PROJECT_DIR / f"과제{hw2}_점검리포트_lms.xlsx"
    build_excel(students_data, out_path)

    print(f"✅ 제출 {len(roster) - n_empty}명 / 미제출 {n_empty}명 / 확인필요 {n_issue}명")
    print(f"✅ xlsx 저장: {out_path}")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="LMS 제출물 내용 확인 (GitHub 미점검 과제용)")
    ap.add_argument("--hw", type=int, default=HW_NUM, help="과제 번호")
    ap.add_argument("--out", type=str, default=None, help="출력 xlsx 경로")
    args = ap.parse_args()
    run(args.hw, args.out)


if __name__ == "__main__":
    main()

"""
import_lms_zip.py
─────────────────
LMS에서 직접 내려받은 '모든 제출물' ZIP을 학생별 폴더로 재편한다.

학교 SSO가 패스키 인증을 요구해 로그인 자동화가 불가능하므로,
ZIP 다운로드까지는 사람이 하고 그 이후는 예전 자동화와 동일하게 처리한다.

사용법:
    1. LMS 과제 → 제출물 목록 → '모든 제출물 내려받기'
    2. 받은 ZIP을 input/ 에 넣는다
    3. python import_lms_zip.py --hw 1
       python import_lms_zip.py --hw 1 --roster "input/스마트팜데이터과학_참여자 목록.xlsx"

결과: lms 제출물/{이름(학번)}/hw{NN}/ 아래에 제출 파일이 정리된다.
"""
from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from pathlib import Path

import openpyxl

from config import (HW_NUM, LMS_DIR, LMS_ZIP_DIR, ROSTER_PATH,
                    SUBMISSION_ROSTER, COL_STUDENT_NAME, COL_STUDENT_ID)
from utils import normalize_name, load_roster_rows, student_dirname

# 참여자 목록 헤더 자동 인식용 (프원실 명단처럼 헤더가 없으면 config 열 번호로 대체)
HEADER_NAME = ("이름", "성명")
HEADER_ID   = ("아이디", "학번", "id")


def load_roster(path: Path) -> dict[str, list[str]]:
    """명단에서 {이름: [학번, ...]} 를 만든다. 동명이인은 학번이 여러 개 담긴다."""
    roster: dict[str, list[str]] = {}
    for name, sid in load_roster_rows(path, COL_STUDENT_NAME, COL_STUDENT_ID):
        roster.setdefault(name, []).append(sid)
    return roster


def find_zip(zip_dir: Path) -> Path:
    """input/ 에서 가장 최근 ZIP 하나를 고른다."""
    zips = sorted(zip_dir.glob("*.zip"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not zips:
        raise FileNotFoundError(
            f"{zip_dir} 에 ZIP 파일이 없습니다. "
            f"LMS에서 '모든 제출물 내려받기'로 받은 ZIP을 넣어 주세요.")
    if len(zips) > 1:
        print(f"  ZIP {len(zips)}개 발견 → 가장 최근 파일 사용: {zips[0].name}")
    return zips[0]


def resolve_folder(raw_name: str, filenames: list[str],
                   roster: dict[str, list[str]]) -> tuple[str, str | None]:
    """
    학생 폴더명을 '이름(학번)' 으로 정한다.

    동명이인은 제출 파일명에 들어 있는 학번으로 구분한다.
    Returns: (폴더명, 경고 메시지 or None)
    """
    ids = roster.get(raw_name, [])

    if len(ids) == 1:
        return student_dirname(raw_name, ids[0]), None

    if len(ids) > 1:                             # 동명이인 → 파일명에서 학번 찾기
        blob = " ".join(filenames)
        hit = [i for i in ids if i and i in blob]
        if len(hit) == 1:
            return student_dirname(raw_name, hit[0]), None
        return raw_name, (f"{raw_name}: 동명이인 {len(ids)}명인데 파일명에서 "
                          f"학번을 특정하지 못해 이름만 사용")

    # 명단에 없음 → 파일명에서 9자리 학번이라도 찾아본다
    m = re.search(r"\b(20\d{7})\b", " ".join(filenames))
    if m:
        return student_dirname(raw_name, m.group(1)), f"{raw_name}: 명단에 없음 (파일명 학번 사용)"
    return raw_name, f"{raw_name}: 명단에 없고 학번도 찾지 못함"


def place_submission(student_folder: Path, dest: Path) -> int:
    """한 학생의 제출 폴더 내용을 dest 아래로 옮긴다. 옮긴 파일 수를 반환."""
    dest.mkdir(parents=True, exist_ok=True)
    moved = 0
    for f in student_folder.iterdir():
        if f.suffix.lower() == ".zip":
            # 학생이 ZIP으로 제출한 경우: 내부 파일을 hw{NN}/ 바로 아래로 펼친다
            with zipfile.ZipFile(f, "r") as zf:
                for member in zf.infolist():
                    if member.is_dir():
                        continue
                    with zf.open(member) as src:
                        (dest / Path(member.filename).name).write_bytes(src.read())
                        moved += 1
        else:
            shutil.move(str(f), str(dest / f.name))
            moved += 1
    return moved


def run(hw_num: int = HW_NUM, zip_path: Path | None = None,
        roster_path: Path | None = None, out_dir: Path | None = None) -> Path:
    """ZIP을 찾아 학생별 폴더로 재편한다."""
    hw2         = f"{hw_num:02d}"
    folder_name = f"hw{hw2}"
    zip_path    = Path(zip_path) if zip_path else find_zip(LMS_ZIP_DIR)
    out_dir     = Path(out_dir) if out_dir else LMS_DIR
    roster      = load_roster(Path(roster_path) if roster_path
                              else (SUBMISSION_ROSTER or ROSTER_PATH))

    print(f"📦 {zip_path.name}")
    n_students = sum(len(v) for v in roster.values())
    print(f"   → {out_dir.name}/[이름(학번)]/{folder_name}/   (명단 {n_students}명)")

    tmp_root = LMS_ZIP_DIR / f"_tmp_{folder_name}"
    if tmp_root.exists():
        shutil.rmtree(tmp_root)
    tmp_root.mkdir(parents=True)

    placed, empty, warns = [], [], []
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(tmp_root)

        # Moodle ZIP이 한 겹 더 감싸여 있는 경우 안쪽으로 내려간다
        work = tmp_root
        roots = [d for d in work.iterdir() if d.is_dir()]
        if len(roots) == 1 and "_assignsubmission_" not in roots[0].name:
            work = roots[0]

        # 1차: 이름 하나뿐이거나, 동명이인이어도 파일명에서 학번이 바로 확정되는 경우
        resolved, pending, used_ids = [], [], {}
        for student_folder in sorted(work.iterdir()):
            if not student_folder.is_dir():
                continue
            raw_name  = student_folder.name.split("_")[0]
            filenames = [f.name for f in student_folder.rglob("*") if f.is_file()]
            ids = roster.get(raw_name, [])

            if len(ids) <= 1:
                dirname, warn = resolve_folder(raw_name, filenames, roster)
                if warn:
                    warns.append(warn)
                resolved.append((student_folder, dirname))
                continue

            blob = " ".join(filenames)
            hit = [i for i in ids if i and i in blob]
            if len(hit) == 1:
                used_ids.setdefault(raw_name, set()).add(hit[0])
                resolved.append((student_folder, student_dirname(raw_name, hit[0])))
            else:
                pending.append((student_folder, raw_name))

        # 2차 소거법: Moodle 제출 ID는 과제마다 바뀌어 재사용할 수 없으므로,
        # "동명이인 중 이미 확정된 사람을 빼면 남는 학번이 하나뿐"인 경우에만 자동 확정한다.
        pending_by_name: dict[str, list[Path]] = {}
        for sf, raw_name in pending:
            pending_by_name.setdefault(raw_name, []).append(sf)

        for raw_name, folders in pending_by_name.items():
            remaining = [i for i in roster.get(raw_name, [])
                        if i not in used_ids.get(raw_name, set())]
            if len(folders) == 1 and len(remaining) == 1:
                resolved.append((folders[0], student_dirname(raw_name, remaining[0])))
                warns.append(f"{raw_name}: 파일명에 학번 없음 — 소거법으로 {remaining[0]} 확정")
            else:
                # 여전히 순서를 알 수 없음 — 한 폴더에 섞이지 않도록 번호를 붙여 분리하고
                # 사람이 직접 확인하도록 알린다.
                for idx, sf in enumerate(folders, 1):
                    resolved.append((sf, f"{raw_name}_확인필요{idx}"))
                warns.append(f"{raw_name}: 동명이인 {len(remaining)}명을 자동으로 구분하지 못함 "
                            f"— {raw_name}_확인필요1~{len(folders)} 폴더를 직접 확인해 이름을 바꿔주세요.")

        for student_folder, dirname in resolved:
            moved = place_submission(student_folder, out_dir / dirname / folder_name)
            (empty if moved == 0 else placed).append(dirname)
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)

    print(f"✅ {len(set(placed))}명 정리 완료")
    if empty:
        print(f"⚠  빈 제출 {len(set(empty))}명: {', '.join(sorted(set(empty)))}")
    for w in warns:
        print(f"⚠  {w}")
    return out_dir


def main():
    ap = argparse.ArgumentParser(description="LMS 제출물 ZIP 재편")
    ap.add_argument("--hw",     type=int, default=HW_NUM, help="과제 번호 (예: 1)")
    ap.add_argument("--zip",    type=str, default=None, help="ZIP 경로 (기본: input/ 최신)")
    ap.add_argument("--roster", type=str, default=None, help="명단 xlsx (기본: config.SUBMISSION_ROSTER)")
    ap.add_argument("--out",    type=str, default=None, help="출력 루트 (기본: config.LMS_DIR)")
    args = ap.parse_args()
    run(args.hw, args.zip, args.roster, args.out)


if __name__ == "__main__":
    main()

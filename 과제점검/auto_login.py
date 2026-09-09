"""
JBNU LMS 과제 제출물 다운로드 스크립트 (반자동).

학교 SSO가 패스키 인증을 요구하므로 로그인은 무인화할 수 없다.
Chrome 프로필을 고정해 세션을 재사용하며, 세션이 없을 때만 브라우저를 띄워
사용자가 직접 패스키 인증을 한다. 인증 이후의 다운로드·분류는 전부 자동이다.
"""

import os
import re
import shutil
import time
import zipfile
from pathlib import Path
from config import (BASE_DIR, OUTPUT_DIR, LMS_DIR, ROSTER_PATH,
                    COL_STUDENT_NAME, CHROME_PROFILE_DIR, LOGIN_WAIT_SEC)
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException


def normalize_name(raw: str) -> str:
    s = str(raw)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[<>:"/\\\\|?*]', "_", s)
    s = s.rstrip(" .")
    return s or "이름없음"

def init_student_dirs(roster_path: Path, student_dir_root: Path, roster_col: int) -> list:
    """
    수강생 명단 xlsx에서 이름을 읽어 학생별 폴더를 사전 생성합니다.

    Returns:
        list[str]: 학생 이름 목록
    """
    wb = openpyxl.load_workbook(roster_path, read_only=True, data_only=True)
    ws = wb.active
    idx = roster_col - 1
    names = [normalize_name(row[idx].value) for row in ws.iter_rows(min_row=2) if row[idx].value]
    wb.close()
    for name in names:
        print(name)
        (student_dir_root / name.split(' ')[0]).mkdir(parents=True, exist_ok=True)
    print(f"학생 폴더 {len(names)}개 준비 완료.")
    return names



LOGIN_URL = "https://lms.jbnu.ac.kr/login/index.php"
LMS_HOME  = "https://lms.jbnu.ac.kr/"


def is_logged_in(driver: webdriver.Chrome) -> bool:
    """Moodle이 body에 붙이는 notloggedin 클래스로 로그인 여부를 판정한다."""
    try:
        cls = driver.find_element(By.TAG_NAME, "body").get_attribute("class") or ""
    except WebDriverException:
        return False
    return "notloggedin" not in cls


def login_to_lms(profile_dir: Path, download_dir: Path,
                 wait_sec: int = 300) -> webdriver.Chrome:
    """
    프로필을 고정한 Chrome으로 LMS에 로그인한다.

    저장된 세션이 살아 있으면 그대로 통과하고, 없으면 JUMP 로그인 화면을 띄운 뒤
    사용자가 패스키 인증을 마칠 때까지 기다린다. (학교 SSO가 패스키를 요구하므로
    이 단계는 자동화할 수 없다 — 인증은 사용자가 직접 수행한다.)

    Args:
        profile_dir  : Chrome 사용자 프로필 경로 (로그인 세션이 여기에 저장된다)
        download_dir : 파일 다운로드 경로
        wait_sec     : 패스키 인증 대기 상한(초)

    Returns:
        로그인이 확인된 드라이버

    Raises:
        RuntimeError: 드라이버를 띄우지 못했거나 대기 시간 내에 로그인되지 않은 경우
    """
    profile_dir.mkdir(parents=True, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={profile_dir.resolve()}")
    options.add_argument("--window-size=1280,900")
    options.add_experimental_option("prefs", {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
    })

    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        options.binary_location = chrome_bin

    try:
        from selenium.webdriver.chrome.service import Service
        chromedriver = os.environ.get("CHROMEDRIVER_PATH")
        service = Service(chromedriver) if chromedriver else None
        driver = webdriver.Chrome(options=options,
                                  **({"service": service} if service else {}))
    except WebDriverException as e:
        raise RuntimeError(f"Chrome WebDriver를 초기화할 수 없습니다: {e}") from e

    try:
        driver.get(LMS_HOME)
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete")

        if is_logged_in(driver):
            print("저장된 세션으로 로그인 상태 확인. 바로 진행합니다.")
            return driver

        # 세션 없음 → 로그인 페이지만 띄우고 사용자가 직접 로그인하도록 기다린다.
        # (버튼 클릭까지 자동화하지 않는다 — SSO가 자동 조작을 거부한다)
        driver.get(LOGIN_URL)

        print("=" * 56)
        print("  브라우저에서 직접 로그인해 주세요.")
        print("  전북대 JUMP 로그인 → 아이디 입력 → 패스키 인증")
        print(f"  최대 {wait_sec}초 대기합니다. (완료되면 자동으로 진행)")
        print("=" * 56)

        deadline = time.time() + wait_sec
        while time.time() < deadline:
            if "lms.jbnu.ac.kr" in driver.current_url and is_logged_in(driver):
                print(f"로그인 확인. (URL: {driver.current_url})")
                return driver
            time.sleep(2)

        raise RuntimeError(f"{wait_sec}초 안에 로그인이 완료되지 않았습니다.")
    except Exception:
        driver.quit()
        raise


def run_task(driver: webdriver.Chrome, course_name: str, download_dir: Path, txt_dir: Path, student_dir_root: Path, student_names: list) -> None:
    """
    로그인 후 지정 강좌/주차의 과제 제출물을 일괄 다운로드합니다.

    Args:
        driver: 로그인된 Chrome 드라이버
        course_name: 강좌명
        download_dir: ZIP 임시 저장 경로
        txt_dir: 과제 설명 txt 저장 경로
        student_dir_root: 학생별 폴더 루트 (lms 제출물/)
        student_names: 수강생 이름 목록
    """
    wait = WebDriverWait(driver, 15)
    target_week = download_dir.name

    # Step 1: LMS 메인 접속
    print(f"[현재 URL] {driver.current_url}")
    print("LMS 메인 페이지 접속 중...")
    driver.get("https://lms.jbnu.ac.kr/")
    print(f"[메인 접속 후 URL] {driver.current_url}")

    # Step 2: 강좌 카드 클릭
    print(f"'{course_name}' 강좌 탐색 중...")
    try:
        course_xpath = f"//ul[contains(@class,'course-lists')]//h5[contains(text(),'{course_name}')]/ancestor::a"
        course_link = wait.until(EC.element_to_be_clickable((By.XPATH, course_xpath)))
        course_link.click()
        print(f"'{course_name}' 강좌 접속 완료.")
    except TimeoutException:
        print(f"[오류] '{course_name}' 강좌를 찾지 못했습니다.")
        return

    # Step 3: 해당 주차 과제 URL 전체 수집
    print(f"'{target_week}' 과제 탐색 중...")
    try:
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "dashboard-ungraded")))
        rows = driver.find_elements(By.CSS_SELECTOR, ".dashboard-ungraded .grid-row")
        target_assignments = []
        for row in rows:
            week_cells = row.find_elements(By.CSS_SELECTOR, ".cell-week span")
            week_title = week_cells[0].get_attribute("data-bs-original-title").strip() if week_cells else ""
            # 완전 일치 대신 포함 여부 + 앞뒤 공백 제거로 비교 (예: " 13주차 ", "제13주차" 등 대응)
            week_num = re.search(r'\d+', target_week)
            title_num = re.search(r'\d+', week_title)
            week_match = (
                week_title == target_week  # 완전 일치 우선
                or week_title.strip() == target_week.strip()  # 공백 제거 일치
                or (week_num and title_num and week_num.group() == title_num.group() and '주차' in week_title)  # 숫자+주차 일치
            )
            if week_cells and week_match:
                activity_links = row.find_elements(By.CSS_SELECTOR, ".cell-activity a")
                if activity_links:
                    target_assignments.append((
                        activity_links[0].text.strip(),
                        activity_links[0].get_attribute("href"),
                    ))

        if not target_assignments:
            # 디버깅: 실제로 LMS에서 감지된 주차 목록 출력
            found_weeks = []
            for row in rows:
                wc = row.find_elements(By.CSS_SELECTOR, ".cell-week span")
                if wc:
                    found_weeks.append(repr(wc[0].get_attribute("data-bs-original-title")))
            print(f"[오류] '{target_week}'에 해당하는 과제를 찾지 못했습니다.")
            print(f"  → dashboard-ungraded에서 감지된 주차 값들: {list(dict.fromkeys(found_weeks))}")
            return

        print(f"총 {len(target_assignments)}개 과제 발견: {[a[0] for a in target_assignments]}")
    except TimeoutException:
        print("[오류] 미채점 활동 목록을 찾지 못했습니다.")
        return

    # Step 4-5: 각 과제별 처리
    for idx, (assignment_name, assignment_url) in enumerate(target_assignments, 1):
        # hw_num 추출 (공통 사용)
        num_match = re.search(r'\d+', assignment_name)
        hw_num = num_match.group().zfill(2) if num_match else str(idx).zfill(2)
        folder_name = f"hw{hw_num}" if num_match else assignment_name

        print(f"\n[{idx}/{len(target_assignments)}] '{assignment_name}' → {folder_name}/ 처리 중...")

        driver.get(assignment_url)

        # 과제 설명 텍스트 저장
        try:
            desc_el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".activity-description")))
            desc_text = desc_el.text.strip()
            txt_path = txt_dir / f"hw{hw_num}.txt"
            txt_path.write_text(desc_text, encoding="utf-8")
            print(f"  과제 설명 저장 완료: {txt_path.name}")
        except TimeoutException:
            print("  [경고] 과제 설명을 찾지 못했습니다.")

        # 제출물 목록 탭
        try:
            submissions_tab = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "li[data-key='mod_assign_submissions'] a")
            ))
            submissions_tab.click()
            print("  제출물 목록 접속 완료.")
        except TimeoutException:
            print("  [오류] '제출물 목록' 탭을 찾지 못했습니다. 건너뜁니다.")
            continue

        # 모든 제출물 내려받기
        try:
            download_btn = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[href*='action=downloadall']")
            ))
            download_btn.click()
            print("  다운로드 시작. 완료 대기 중...")

            for _ in range(60):
                time.sleep(1)
                if not list(download_dir.glob("*.crdownload")):
                    break

            # 다운로드된 ZIP → hw{num}.zip 으로 rename (스킵 마커 겸 보관)
            zip_files = sorted(download_dir.glob("*.zip"), key=lambda f: f.stat().st_mtime, reverse=True)
            if zip_files:
                hw_zip = download_dir / f"hw{hw_num}.zip"
                zip_files[0].replace(hw_zip)

                # 학생별 폴더로 재편
                tmp_dir = download_dir / f"_tmp_hw{hw_num}"
                tmp_dir.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(hw_zip, 'r') as zf:
                    zf.extractall(tmp_dir)

                for student_folder in tmp_dir.iterdir():
                    if not student_folder.is_dir():
                        continue
                    raw_name = student_folder.name.split('_')[0]
                    matched = next((n for n in student_names if n == raw_name), raw_name)
                    dest = student_dir_root / matched / folder_name
                    dest.mkdir(parents=True, exist_ok=True)

                    for f in student_folder.iterdir():
                        if f.suffix.lower() == '.zip':
                            # 학생 제출 zip: 내부 파일을 hw{num}/ 바로 아래에 flatten
                            with zipfile.ZipFile(f, 'r') as zf:
                                for member in zf.infolist():
                                    if member.is_dir():
                                        continue
                                    filename = Path(member.filename).name
                                    with zf.open(member) as src:
                                        (dest / filename).write_bytes(src.read())
                        else:
                            shutil.move(str(f), str(dest / f.name))

                shutil.rmtree(tmp_dir)
                print(f"  재편 완료: [학생명]/{folder_name}/")
            else:
                print("  [경고] ZIP 파일을 찾지 못했습니다.")

        except TimeoutException:
            print("  [오류] '모든 제출물 내려받기' 버튼을 찾지 못했습니다.")

    print(f"\n전체 완료.")
    print(f"  제출물: {student_dir_root}")
    print(f"  과제설명: {txt_dir}")


def main(TARGET_COURSE, TARGET_WEEK) -> None:
    """
    LMS 제출물 다운로드 진입점.

    로그인은 반자동이다 — 저장된 세션이 없으면 브라우저가 열리고,
    사용자가 패스키 인증을 마친 뒤 자동으로 다운로드가 진행된다.
    """
    DOWNLOAD_DIR = OUTPUT_DIR / TARGET_WEEK              # ZIP 임시 저장
    TXT_DIR      = BASE_DIR / '과제모음'                  # 과제 설명 txt
    STUDENT_DIR  = LMS_DIR                               # 학생별 폴더 루트

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    TXT_DIR.mkdir(parents=True, exist_ok=True)

    student_names = init_student_dirs(ROSTER_PATH, STUDENT_DIR, COL_STUDENT_NAME)

    driver = login_to_lms(CHROME_PROFILE_DIR, DOWNLOAD_DIR, LOGIN_WAIT_SEC)
    try:
        run_task(driver, TARGET_COURSE, DOWNLOAD_DIR, TXT_DIR,
                 STUDENT_DIR, student_names)
    finally:
        driver.quit()


if __name__ == "__main__":
    from config import TARGET_COURSE, TARGET_WEEK
    main(TARGET_COURSE, TARGET_WEEK)

"""
JBNU LMS 자동 로그인 및 과제 제출물 다운로드 스크립트.

이 스크립트는 Selenium과 python-dotenv를 활용하여
.env 파일에 저장된 사용자 자격 증명으로 LMS에 자동 로그인 후
지정한 주차의 과제 제출물을 일괄 다운로드합니다.
"""

import os
import re
import shutil
import time
import zipfile
from pathlib import Path
from dotenv import load_dotenv
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys


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



def login_to_lms(user_id: str, user_pw: str, otp_secret: str = None, options: webdriver.ChromeOptions = None) -> webdriver.Chrome:
    """
    제공된 아이디와 비밀번호로 JBNU LMS 홈페이지에 로그인합니다.

    Returns:
        webdriver.Chrome: 로그인 완료된 브라우저 드라이버 객체
    """
    login_url = "https://lms.jbnu.ac.kr/login/index.php"

    if options is None:
        options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        options.binary_location = chrome_bin

    try:
        from selenium.webdriver.chrome.service import Service
        chromedriver = os.environ.get("CHROMEDRIVER_PATH")
        service = Service(chromedriver) if chromedriver else None
        driver = webdriver.Chrome(options=options, **({"service": service} if service else {}))
    except WebDriverException as e:
        print(f"[오류] Chrome WebDriver를 초기화할 수 없습니다: {e}")
        return None

    try:
        print(f"[{login_url}] 접속 중...")
        driver.get(login_url)

        wait = WebDriverWait(driver, 10)

        print("아이디와 비밀번호 입력 중...")
        id_input = wait.until(EC.presence_of_element_located((By.ID, "input-username")))
        id_input.clear()
        id_input.send_keys(user_id)

        pw_input = wait.until(EC.presence_of_element_located((By.ID, "input-password")))
        pw_input.clear()
        pw_input.send_keys(user_pw)

        print("로그인 버튼 클릭 시도 (엔터키 전송)...")
        pw_input.send_keys(Keys.RETURN)
        print("로그인 동작이 완료되었습니다.")

        if otp_secret:
            print("OTP 인증 설정이 확인되었습니다. 인증번호 생성을 시도합니다...")
            import onetimepass as otp

            # 서버 시계 오차 보정: NTP로 정확한 UTC 시각 획득
            try:
                import ntplib
                ntp_time = ntplib.NTPClient().request('pool.ntp.org', version=3).tx_time
                print(f"NTP 시각 동기화 성공: {ntp_time}")
            except Exception as e:
                import time
                ntp_time = time.time()
                print(f"NTP 동기화 실패 ({e}), 시스템 시각 사용")

            my_token = str(otp.get_totp(otp_secret, clock=int(ntp_time))).zfill(6)
            print(f"생성된 OTP 인증번호: {my_token} 입력 중...")

            try:
                otp_xpath = "/html/body/div[2]/div[2]/div/div/section/div/div[2]/form/div[2]/div[2]/input"
                time.sleep(3)

                success = False
                for attempt in range(5):
                    try:
                        # OTP 만료 방지: 매 시도마다 토큰 재생성
                        my_token = str(otp.get_totp(otp_secret)).zfill(6)
                        otp_input = wait.until(EC.presence_of_element_located((By.XPATH, otp_xpath)))
                        # JS로 값 주입 + 이벤트 발생 (stale 방지)
                        driver.execute_script("""
                            arguments[0].value = arguments[1];
                            arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                            arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
                        """, otp_input, my_token)
                        # submit 버튼 클릭 (Enter 대신)
                        try:
                            submit_btn = driver.find_element(By.CSS_SELECTOR, "form button[type='submit'], form input[type='submit']")
                            submit_btn.click()
                        except Exception:
                            driver.execute_script("arguments[0].form.submit()", otp_input)
                        # URL이 mfa 페이지를 벗어날 때까지 대기
                        from selenium.webdriver.support.ui import WebDriverWait as WDW
                        try:
                            WDW(driver, 10).until(lambda d: "mfa" not in d.current_url)
                            print(f"OTP 인증 성공. 현재 URL: {driver.current_url}")
                        except Exception:
                            print(f"[경고] OTP 제출 후 페이지 미전환. 현재 URL: {driver.current_url}")
                        success = True
                        break
                    except StaleElementReferenceException:
                        print(f"Stale 오류, 재시도 중... ({attempt + 1}/5)")
                        time.sleep(1)
                if success:
                    print("OTP 인증번호 전송 완료.")
                else:
                    print("[경고] OTP 입력 5회 모두 실패.")
            except TimeoutException:
                print("OTP 입력창을 찾지 못했습니다.")
            except Exception as e:
                print(f"OTP 처리 중 기타 오류 발생: {e}")

        time.sleep(3)

    except TimeoutException:
        print("[오류] 페이지 로딩 또는 요소를 찾는데 시간이 초과되었습니다.")
    except Exception as e:
        print(f"[오류] 로그인 처리 중 알 수 없는 오류 발생: {e}")
    finally:
        print("로그인 프로세스 종료")
        return driver


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
# def main() -> None:

    """
    메인 실행 함수.
    """
    # ────── 사용자 설정 ──────
    TARGET_COURSE = TARGET_COURSE
    TARGET_WEEK   = TARGET_WEEK
    DOWNLOAD_DIR  = Path('.') / 'output' / TARGET_WEEK                        # ZIP 임시 저장
    TXT_DIR       = Path('..') / '과제점검' / '과제모음'                       # 과제 설명 txt
    STUDENT_DIR    = Path('..') / 'lms 제출물'                                 # 학생별 폴더 루트
    ROSTER_PATH    = Path('..') / '과제점검' / '프원실(2026)_과제확인.xlsx'      # 수강생 명단
    ROSTER_COL     = 5                                                      # 학생 영문 이름(1-based)
    # ─────────────────────────

    env_path = Path('.') / '.env'
    load_dotenv(dotenv_path=env_path)

    user_id = os.environ.get("LMS_USER_ID")
    user_pw = os.environ.get("LMS_USER_PW")
    otp_secret = os.environ.get("LMS_OTP_SECRET")

    if not user_id or not user_pw:
        print("[오류] .env 파일에 LMS_USER_ID 또는 LMS_USER_PW가 설정되지 않았습니다.")
        return

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    TXT_DIR.mkdir(parents=True, exist_ok=True)

    student_names = init_student_dirs(ROSTER_PATH, STUDENT_DIR, ROSTER_COL)

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
    })

    driver = login_to_lms(user_id, user_pw, otp_secret, options=options)

    if driver:
        run_task(driver, TARGET_COURSE, DOWNLOAD_DIR, TXT_DIR, STUDENT_DIR, student_names)


if __name__ == "__main__":
    main("프로그래밍원리와실습", '16주차')


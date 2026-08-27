"""
TA 근무일지 HWPX 표 내용 교체 스크립트 — 4월판
=========================================
사용법:
    python update_ta_근무일지_4월.py

입력:  TA_근무일지_양식.hwpx  (스크립트와 같은 폴더)
출력:  TA_근무일지_4월.hwpx  (같은 폴더에 생성)

특징:
  - 표준 라이브러리(zipfile, os, re)만 사용 — pip 설치 불필요
  - 원본 행을 그대로 복제 후 텍스트만 교체 → 서식/테두리/컨트롤 보존
  - 행 수가 늘어나도 자동 처리
  - NEW_ROWS 리스트와 TOTAL_HOURS만 수정하면 다음 달에도 재사용 가능
"""

import zipfile
import re
import os

# ─────────────────────────────────────────────────────────
#  경로 설정
# ─────────────────────────────────────────────────────────
INPUT_HWPX  = "TA_근무일지_양식.hwpx"
OUTPUT_HWPX = "TA_근무일지_4월.hwpx"

# ─────────────────────────────────────────────────────────
#  ★ 4월 근무 데이터 ★
#
#  일정 구성:
#    화요일(09:00~11:00): 4/7, 14, 21, 28  → 4회 × 2h = 8h
#    목요일(11:00~13:00): 4/2, 9, 16, 23, 30 → 5회 × 2h = 10h
#    추가(월/수/금):     4/4, 8, 13, 17, 22, 29 → 6회 × 2h = 12h
#    ─────────────────────────────────────────────────────
#    합계: 15회 × 2h = 30h
#
#  과제 내용 반영:
#    HW07: 칼로리 계산 + 마트 예제 (반복문, 사전)
#    HW08: 메인 함수 도입, gugudan/c2f/sum_n 함수
#    HW09: 함수 모듈화, average/get_range_list/is_leap_year
#    HW10: 파일 읽기, 통계(개수·평균·최댓값·최솟값·중앙값)
#    HW11: CSV 기상자료, 연평균기온·강우일수·총강우량
#    HW12: 기상자료 활용 5개 프로그램
#
#  형식: (근무일자, 요일, 시작시각, 종료시각, 근무시간, 내용1줄, 내용2줄)
#  내용이 한 줄이면 내용2줄에 빈 문자열 "" 입력
# ─────────────────────────────────────────────────────────
NEW_ROWS = [
    # ── HW07: 칼로리 계산 + 마트 예제 (반복문, 사전) ──────────────────
    ("26. 4. 2.",  "목", "11:00", "13:00", "2", "HW07 칼로리 계산 및 마트 예제 수업 보조 진행", "반복문·사전 활용 데이터 처리 구조 실습 지도"),
    ("26. 4. 4.",  "금", "14:00", "16:00", "2", "HW07 제출물 전체 검토 및 피드백 작성 진행", "사전 키 오류, 반복문 조건 누락 사례 중점 점검"),
    ("26. 4. 7.",  "화", "09:00", "11:00", "2", "HW07 오류 사항 수업 보조 및 질의응답 진행", "마트 예제 분기 처리, 칼로리 합산 오류 지도"),
    ("26. 4. 8.",  "수", "14:00", "16:00", "2", "HW07 검토 완료 후 이슈 목록 작성 및 정리", "LMS 판정 기준 검토 및 제출 현황 최종 확인"),
    ("26. 4. 9.",  "목", "11:00", "13:00", "2", "HW07 오류 개별 피드백 및 재제출 안내 진행", "dict 키 오류, 반복 종료 조건 미설정 개별 상담"),

    # ── HW08: 메인 함수 도입, gugudan/c2f/sum_n ──────────────────────
    ("26. 4. 13.", "월", "14:00", "16:00", "2", "HW08 메인 함수 도입 수업 보조 및 안내 진행", "main() 구조, __name__ 조건문 사용법 실습 지도"),
    ("26. 4. 14.", "화", "09:00", "11:00", "2", "HW08 함수 작성 수업 보조 및 질의응답 진행", "gugudan·c2f·sum_n 함수 구현 및 테스트 지도"),
    ("26. 4. 16.", "목", "11:00", "13:00", "2", "HW08 오류 사항 수업 보조 및 개별 피드백", "반환값 오류, 함수 호출 위치 문제 개별 상담"),
    ("26. 4. 17.", "금", "14:00", "16:00", "2", "HW08 검토 완료 후 이슈 목록 작성 및 정리", "함수 정의 위치, main 조건 미적용 사례 정리"),

    # ── HW09: 함수 모듈화, average/get_range_list/is_leap_year ─────────
    ("26. 4. 21.", "화", "09:00", "11:00", "2", "HW09 함수 모듈화 수업 보조 및 질의응답 진행", "average·get_range_list·is_leap_year 구현 지도"),
    ("26. 4. 22.", "수", "14:00", "16:00", "2", "HW09 제출물 전체 검토 및 피드백 작성 진행", "윤년 조건 오류, 리스트 반환 누락 사례 점검"),
    ("26. 4. 23.", "목", "11:00", "13:00", "2", "HW09 오류 사항 수업 보조 및 개별 피드백", "import 시 __name__ 미적용 실행 오류 개별 상담"),

    # ── HW10·HW11·HW12: 파일 읽기, 통계, 기상자료 ───────────────────
    ("26. 4. 28.", "화", "09:00", "11:00", "2", "HW10 파일 읽기 및 통계 함수 수업 보조 진행", "개수·평균·최댓값·최솟값·중앙값 출력 실습 지도"),
    ("26. 4. 29.", "수", "14:00", "16:00", "2", "HW10·HW11 제출물 전체 검토 및 피드백 작성", "파일 파싱 오류, CSV 기상자료 처리 방법 점검"),
    ("26. 4. 30.", "목", "11:00", "13:00", "2", "HW11·HW12 기상자료 분석 수업 보조 진행", "연평균기온·강우일수·5개 프로그램 실습 지도"),
]

TOTAL_HOURS = "30"  # 총 근무시간 (숫자만)


# ─────────────────────────────────────────────────────────
#  원본 파일에서 데이터 행이 시작되는 tr 인덱스 (변경 금지)
#  tr[0]~tr[13]: 헤더/인적사항/열제목 행
#  tr[14]~tr[25]: 근무 데이터 행 12개
#  tr[26]: 총 근무시간 합계 행
# ─────────────────────────────────────────────────────────
DATA_START_IDX      = 14
ORIG_DATA_ROW_COUNT = 12


# ─────────────────────────────────────────────────────────
#  유틸 함수
# ─────────────────────────────────────────────────────────

def esc(text):
    """XML 특수문자 이스케이프"""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))


def split_tcs(tr_xml):
    """tr 안의 tc 블록들을 리스트로 분리"""
    return re.findall(r'<hp:tc\b.*?</hp:tc>', tr_xml, re.DOTALL)


def replace_hp_t(tc_xml, new_text, occurrence=1):
    """tc 블록 안의 hp:t 를 occurrence 번째만 교체 (나머지 보존)"""
    count = [0]
    def repl(m):
        count[0] += 1
        return f'<hp:t>{esc(new_text)}</hp:t>' if count[0] == occurrence else m.group(0)
    return re.sub(r'<hp:t>[^<]*</hp:t>', repl, tc_xml)


def make_row(template_tr, row_data, row_idx):
    """
    원본 tr 하나를 복제해 row_data 내용으로 교체한 새 tr XML 반환.
    서식/테두리/컨트롤(서명란 fieldBegin 등)은 그대로 유지된다.
    """
    date, day, start, end, hours, line1, line2 = row_data
    new_row_addr = DATA_START_IDX + row_idx

    tcs = split_tcs(template_tr)
    # tc 순서: [0]날짜, [1]요일, [2]시작, [3]~, [4]종료, [5]시간, [6]내용, [7]서명
    tcs[0] = replace_hp_t(tcs[0], date)
    tcs[1] = replace_hp_t(tcs[1], day)
    tcs[2] = replace_hp_t(tcs[2], start)
    # tcs[3] = '~' 고정
    tcs[4] = replace_hp_t(tcs[4], end)
    tcs[5] = replace_hp_t(tcs[5], hours)
    tcs[6] = replace_hp_t(tcs[6], line1, occurrence=1)
    if line2:
        tcs[6] = replace_hp_t(tcs[6], line2, occurrence=2)
    # tcs[7] = 서명란 고정

    # rowAddr 업데이트
    tcs = [re.sub(r'rowAddr="\d+"', f'rowAddr="{new_row_addr}"', tc) for tc in tcs]

    # 서명란 fieldBegin/beginIDRef id 업데이트 (중복 방지)
    new_fid = 2123213024 + row_idx
    tcs[7] = re.sub(r'fieldBegin id="\d+"', f'fieldBegin id="{new_fid}"', tcs[7])
    tcs[7] = re.sub(r'beginIDRef="\d+"',    f'beginIDRef="{new_fid}"',    tcs[7])

    return '<hp:tr>' + ''.join(tcs) + '</hp:tr>'


# ─────────────────────────────────────────────────────────
#  메인 처리
# ─────────────────────────────────────────────────────────

def main():
    if not os.path.exists(INPUT_HWPX):
        print(f"[오류] 입력 파일 없음: {INPUT_HWPX}")
        print("       이 스크립트와 같은 폴더에 HWPX 파일을 두세요.")
        return

    # 원본 section0.xml 읽기
    with zipfile.ZipFile(INPUT_HWPX) as z:
        xml = z.read('Contents/section0.xml').decode('utf-8')

    # tr 블록 전체 파악
    TR_PAT = re.compile(r'<hp:tr>.*?</hp:tr>', re.DOTALL)
    all_matches = list(TR_PAT.finditer(xml))
    total_orig = len(all_matches)
    data_end_idx = DATA_START_IDX + ORIG_DATA_ROW_COUNT

    print(f"[정보] 원본 전체 행: {total_orig}  /  데이터 행: {ORIG_DATA_ROW_COUNT}  →  새 데이터 행: {len(NEW_ROWS)}")

    # 템플릿: 첫 번째 데이터 행(tr[14])을 기준으로 복제
    template_tr = all_matches[DATA_START_IDX].group()

    # 새 데이터 행 생성
    new_data_rows = [make_row(template_tr, row, i) for i, row in enumerate(NEW_ROWS)]

    # XML 재조립: 데이터 행 구간을 통째로 교체
    data_block_start = all_matches[DATA_START_IDX].start()
    data_block_end   = all_matches[data_end_idx - 1].end()

    suffix = xml[data_block_end:]

    # 뒷부분 행들의 rowAddr를 행 수 변화만큼 이동
    delta = len(NEW_ROWS) - ORIG_DATA_ROW_COUNT
    if delta != 0:
        def shift_addr(m):
            v = int(m.group(1))
            return f'rowAddr="{v + delta}"' if v >= data_end_idx else m.group(0)
        suffix = re.sub(r'rowAddr="(\d+)"', shift_addr, suffix)

    new_xml = xml[:data_block_start] + ''.join(new_data_rows) + suffix

    # rowCnt 업데이트
    new_xml = re.sub(r'rowCnt="\d+"', f'rowCnt="{total_orig + delta}"', new_xml, count=1)

    # 총 근무시간 교체 (합계 행의 숫자 셀 — 첫 번째 등장만)
    new_xml = new_xml.replace('<hp:t>30</hp:t>', f'<hp:t>{TOTAL_HOURS}</hp:t>', 1)

    # 새 HWPX 저장 (section0.xml만 교체, 나머지는 원본 그대로)
    tmp = OUTPUT_HWPX + '.tmp'
    with zipfile.ZipFile(INPUT_HWPX) as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                item.flag_bits |= 0x800  # UTF-8 플래그 명시적 설정
                if item.filename == 'Contents/section0.xml':
                    data = new_xml.encode('utf-8')
                    # XML 선언에 encoding="UTF-8" 보장
                    if not data.startswith(b'<?xml'):
                        data = b'<?xml version="1.0" encoding="UTF-8"?>\n' + data
                    zout.writestr(item, data)
                else:
                    zout.writestr(item, zin.read(item.filename))
    if os.path.exists(OUTPUT_HWPX):
        os.remove(OUTPUT_HWPX)
    os.rename(tmp, OUTPUT_HWPX)

    print(f"[완료] {OUTPUT_HWPX} 생성됨")
    print(f"       총 {len(NEW_ROWS)}개 행, 총 근무시간 {TOTAL_HOURS}시간")


if __name__ == "__main__":
    main()

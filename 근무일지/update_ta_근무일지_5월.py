"""
TA 근무일지 HWPX 표 내용 교체 스크립트 — 5월판
=========================================
사용법:
    python update_ta_근무일지_5월.py

입력:  TA_근무일지_양식.hwpx  (스크립트와 같은 폴더)
출력:  TA_근무일지_5월.hwpx  (같은 폴더에 생성)
"""

import zipfile
import re
import os

# ─────────────────────────────────────────────────────────
#  경로 설정
# ─────────────────────────────────────────────────────────
INPUT_HWPX  = "TA_근무일지_양식.hwpx"
OUTPUT_HWPX = "TA_근무일지_5월.hwpx"

# ─────────────────────────────────────────────────────────
#  ★ 5월 근무 데이터 ★
#
#  일정 구성:
#    공휴일 제외: 5/1(근로자의날), 5/5(어린이날)
#    출장 제외: 5/25~5/31
#    → 가용 평일 14일 중 선택
#    5/15(금), 5/22(금) 3시간 근무로 시간 보충
#    ─────────────────────────────────────────────────────
#    합계: 12회×2h + 2회×3h = 30h
#
#  과제 내용 반영:
#    HW13: 5페이지 1·2번 — 다년간(2001~2022) 기상자료 분석
#    HW14: 최대일교차 날짜·값, 5~9월 적산온도 (2001~2022)
#    HW15: 2023년 기상자료 — 연평균기온·강우일수·총강우량
#    HW16: 연구실 안전교육 수료증 제출
#    HW17: 숫자 리스트 입력·출력, 총 개수·평균
#
#  형식: (근무일자, 요일, 시작시각, 종료시각, 근무시간, 내용1줄, 내용2줄)
#  내용이 한 줄이면 내용2줄에 빈 문자열 "" 입력
# ─────────────────────────────────────────────────────────
NEW_ROWS = [
    # ── HW13: 다년간 기상자료 분석 (2001~2022, 5페이지 1·2번) ──────────
    ("26. 5. 4.",  "월", "14:00", "16:00", "2", "HW13 기상자료 분석 수업 보조 및 안내 진행", "2001~2022년 다년간 데이터 처리 방법 실습 지도"),
    ("26. 5. 6.",  "수", "14:00", "16:00", "2", "HW13 기상자료 프로그램 수업 보조 및 질의응답", "CSV 파일 읽기 오류, 다중 연도 처리 구조 지도"),
    ("26. 5. 7.",  "목", "11:00", "13:00", "2", "HW13 제출물 전체 검토 및 피드백 작성 진행", "연도별 통계 계산 오류, 데이터 파싱 문제 점검"),

    # ── HW14: 최대일교차·적산온도 (2001~2022) ─────────────────────────
    ("26. 5. 11.", "월", "14:00", "16:00", "2", "HW14 최대일교차 분석 수업 보조 및 안내 진행", "연도별 최대일교차 날짜·값 출력 구조 실습 지도"),
    ("26. 5. 12.", "화", "09:00", "11:00", "2", "HW14 적산온도 계산 수업 보조 및 질의응답", "5~9월 적산온도 산정 방법 및 반복 구조 지도"),
    ("26. 5. 13.", "수", "14:00", "16:00", "2", "HW14 제출물 전체 검토 및 피드백 작성 진행", "일교차 계산 오류, 적산온도 조건 누락 사례 점검"),
    ("26. 5. 14.", "목", "11:00", "13:00", "2", "HW14 오류 사항 수업 보조 및 개별 피드백", "날짜 파싱 오류, 월별 필터링 조건 미적용 상담"),

    # ── HW15: 2023년 기상자료 분석 ────────────────────────────────────
    ("26. 5. 18.", "월", "14:00", "16:00", "2", "HW15 2023년 기상자료 수업 보조 및 안내 진행", "연 평균 기온·강우일수·총 강우량 계산 실습 지도"),
    ("26. 5. 19.", "화", "09:00", "11:00", "2", "HW15 기상자료 분석 수업 보조 및 질의응답", "2023년 CSV 파일 구조 파악 및 통계 처리 지도"),
    ("26. 5. 20.", "수", "14:00", "16:00", "2", "HW15 제출물 전체 검토 및 피드백 작성 진행", "파일 인코딩 오류, 조건 필터링 누락 사례 점검"),

    # ── HW16: 연구실 안전교육 수료증 ──────────────────────────────────
    ("26. 5. 21.", "목", "11:00", "13:00", "2", "HW16 연구실 안전교육 수료증 제출 확인 진행", "수료증 미제출 학생 개별 안내 및 제출 독려 진행"),

    # ── HW17: 숫자 리스트 입력·출력 (총 개수·평균) ────────────────────
    ("26. 5. 8.",  "금", "14:00", "16:00", "2", "HW17 숫자 리스트 입력 수업 보조 및 안내 진행", "정수 입력 필터링, -1 종료 조건 구현 실습 지도"),
    ("26. 5. 15.", "금", "14:00", "17:00", "3", "HW17 수업 보조 및 질의응답 진행", "총 개수·평균 계산, 자연수 필터링 로직 지도"),
    ("26. 5. 22.", "금", "14:00", "17:00", "3", "HW17 제출물 전체 검토 및 오류 개별 피드백", "음수 입력 처리, -1 조건 오류 사례 개별 상담"),
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

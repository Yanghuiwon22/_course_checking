"""
TA 근무일지 HWPX 표 내용 교체 스크립트 — 6월판
=========================================
사용법:
    python update_ta_근무일지_6월.py

입력:  TA_근무일지_양식.hwpx  (스크립트와 같은 폴더)
출력:  TA_근무일지_6월.hwpx  (같은 폴더에 생성)
"""

import zipfile
import re
import os

# ─────────────────────────────────────────────────────────
#  경로 설정
# ─────────────────────────────────────────────────────────
INPUT_HWPX  = "TA_근무일지_양식.hwpx"
OUTPUT_HWPX = "TA_근무일지_6월.hwpx"

# ─────────────────────────────────────────────────────────
#  ★ 6월 근무 데이터 ★
#
#  근무기간: 2026. 6. 1.(월) ~ 6. 19.(금)  ※이후 활동불가
#  공휴일 없음 (현충일 6/6은 토요일)
#  → 가용 평일 15일 전체 사용, 1일 2시간 근무
#  ─────────────────────────────────────────────────────
#  합계: 15회×2h = 30h
#
#  과제 내용 반영:
#    HW18: 대소문자 변환(toggle_text), 카이사르 암호, 초성 게임
#    HW19: 카운트다운, 구구단 문제집, 행맨, 로또번호추출기 (4개 프로그램)
#    HW20: 전주시·수원시 기상자료 분석 (강수량·최대기온·일교차·그래프)
#    HW21: HW19 GUI 변환(gui_input/rich/PySimpleGUI), GitHub 업로드
#    기말과제점검: Docker·Streamlit 실행 환경 점검 및 최종 피드백
#
#  형식: (근무일자, 요일, 시작시각, 종료시각, 근무시간, 내용1줄, 내용2줄)
#  내용이 한 줄이면 내용2줄에 빈 문자열 "" 입력
# ─────────────────────────────────────────────────────────
NEW_ROWS = [
    # ── HW18: 대소문자 변환·카이사르 암호·초성 게임 ────────────────────
    ("26. 6. 1.",  "월", "14:00", "16:00", "2", "HW18 대소문자 변환·카이사르 암호 수업 보조 및 안내 진행", "ASCII 코드 활용 toggle_text, caesar_encode/decode 구현 실습 지도"),
    ("26. 6. 2.",  "화", "09:00", "11:00", "2", "HW18 초성 게임 구현 수업 보조 및 질의응답", "유니코드 활용 초성 추출, 단어 매칭 로직 지도"),
    ("26. 6. 3.",  "수", "14:00", "16:00", "2", "HW18 제출물 전체 검토 및 피드백 작성 진행", "ASCII 변환 오류, 카이사르 암호 시프트 처리 오류 점검"),

    # ── HW19: 카운트다운·구구단 문제집·행맨·로또번호추출기 (4개) ───────
    ("26. 6. 4.",  "목", "11:00", "13:00", "2", "HW19 카운트다운·구구단 문제집 수업 보조 및 안내 진행", "타이머 구현, 구구단 랜덤 출제·정답 채점 로직 실습 지도"),
    ("26. 6. 5.",  "금", "14:00", "16:00", "2", "HW19 행맨 게임 수업 보조 및 질의응답", "단어 추측, 남은 시도 횟수 처리 구조 지도"),
    ("26. 6. 8.",  "월", "14:00", "16:00", "2", "HW19 로또번호추출기 수업 보조 및 안내 진행", "중복 없는 난수 추출, 정렬 출력 구현 지도"),
    ("26. 6. 9.",  "화", "09:00", "11:00", "2", "HW19 제출물 전체 검토 및 피드백 작성 진행", "4개 프로그램 누락 사례, 반복문 구조 오류 점검"),

    # ── HW20: 전주시·수원시 기상자료 분석 ───────────────────────────────
    ("26. 6. 10.", "수", "14:00", "16:00", "2", "HW20 기상자료 분석 수업 보조 및 안내 진행", "전주시·수원시 연 강수량, 최대기온·일교차 계산 실습 지도"),
    ("26. 6. 11.", "목", "11:00", "13:00", "2", "HW20 선그래프·막대그래프 시각화 수업 보조 및 질의응답", "matplotlib 활용 연도별 평균기온·강수량 그래프 작성 지도"),
    ("26. 6. 12.", "금", "14:00", "16:00", "2", "HW20 생일 기준 온도 분석 수업 보조 및 질의응답", "지정 일자 기준 연도별 온도 비교, 최고·최저 연도 산출 지도"),
    ("26. 6. 15.", "월", "14:00", "16:00", "2", "HW20 제출물 전체 검토 및 피드백 작성 진행", "데이터 필터링 오류, 그래프 축 설정 누락 사례 점검"),

    # ── HW21: HW19 GUI 변환 ─────────────────────────────────────────────
    ("26. 6. 16.", "화", "09:00", "11:00", "2", "HW21 GUI 변환 수업 보조 및 안내 진행", "input()을 gui_input()/rich/PySimpleGUI로 대체하는 방법 실습 지도"),
    ("26. 6. 17.", "수", "14:00", "16:00", "2", "HW21 제출물 전체 검토 및 피드백 작성 진행", "GUI 미적용 사례, GitHub 업로드 누락 학생 개별 안내"),

    # ── 기말과제점검 ─────────────────────────────────────────────────────
    ("26. 6. 18.", "목", "11:00", "13:00", "2", "기말과제 점검 환경 구축 및 수업 보조 진행", "Docker·Streamlit 실행 환경 점검, 제출 코드 구조 확인"),
    ("26. 6. 19.", "금", "14:00", "16:00", "2", "기말과제 제출물 전체 점검 및 최종 피드백 작성 진행", "학생별 실행 오류, 기능 누락 사항 정리 및 안내"),
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

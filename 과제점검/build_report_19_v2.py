"""
build_report_19_v2.py
과제19 점검 리포트 — hw18 스타일 참고
시트: hw19 점검결과 (단일 시트)
열: 이름/학번/GitHub ID / LMS/GitHub/로직이슈/최종종합 / [빈칸] / 코드체크4개 / [빈칸] / 점수4개 / [빈칸] / 우수 / 최종
"""
from datetime import date
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

TODAY = date.today().strftime("%Y-%m-%d")
OUT_PATH = Path(__file__).parent.parent / "과제19_점검리포트.xlsx"

# ── 색상 ──────────────────────────────────────────────────────────
GRAY   = "D9D9D9"   # 미제출
WHITE  = "FFFFFF"

# ── 스타일 헬퍼 ──────────────────────────────────────────────────
_thin = Side(border_style="thin", color="D0D0D0")

def _border():
    return Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

def _fill(hex_):
    return PatternFill("solid", fgColor=hex_)

def _font(bold=False, size=9, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)

def _align(h="center", wrap=True):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def wcell(ws, r, c, val, *, bold=False, sz=9, fg="000000", bg=None,
          h="center", wrap=True, italic=False):
    cl = ws.cell(r, c, val)
    cl.font      = _font(bold=bold, size=sz, color=fg, italic=italic)
    cl.alignment = _align(h, wrap)
    cl.border    = _border()
    if bg:
        cl.fill  = _fill(bg)
    return cl

# ── 학생 데이터 ───────────────────────────────────────────────────
# (이름, 학번, GitHub_ID, LMS, GitHub제출, 로직이슈, 최종종합,
#  H행맨, I구구단, J로또, K대체구현)
STUDENTS = [
    ("강태현",       "202515406", "kth0208",
     "-",  "이슈",
     "[gugudan1.py] 점수 출력이 for 루프 안에서 매번 출력됨",
     "이슈", False, True, False, False),

    ("김근영",       "202515408", "zbr2s5rjc9-sys",
     "-",  "경미한 이슈",
     "[countdown.py] 사용자 입력 없이 5부터 고정 카운트 (minor); [lotto.py] 로또 로직이 main() 밖에서 실행됨 (minor)",
     "경미한 이슈", False, False, False, False),

    ("김도영",       "202210033", "Doyoung03",
     "미제출", "이슈",
     "[hw3.py] f-string 중첩으로 SyntaxError 발생 / 맞혀도 trial 감소하는 행맨 로직 오류",
     "LMS 미제출", True, False, False, False),

    ("김민성",       "202321610", "Kim-cloud-ai",
     "-",  "이슈복수",
     "[homework19-2.py] count+=0으로 틀린 횟수 집계 안 됨; [homework19-3.py] 행맨 대신 약어 맞추기 구현 (행맨 미구현); [homework19-4.py] 로또 중복 번호 허용",
     "이슈복수", False, True, True, True),

    ("김선우",       "202515414", "SeonU777",
     "-",  "이슈",
     '[hang_man.py] w_list[count] 순서 비교로 행맨 규칙 미준수; [count_down.py] end="" 로 숫자 이어붙여 출력 (minor)',
     "이슈", True, False, False, False),

    ("김선호",       "202515415", "fnfnf2145",
     "미제출", "이상없음",
     "이상없음",
     "LMS 미제출", False, False, False, False),

    ("김예지",       "202246491", "dpwl829",
     "미제출", "미제출",
     "—",
     "미제출", False, False, False, False),

    ("김준혁",       "202310054", "kimjunhyeok04",
     "-",  "이상없음",
     "[HW19] .py 없이 ZIP 파일로만 제출 (minor)",
     "이상없음", False, False, False, False),

    ("김현준",       "202310055", "202310055HJ",
     "-",  "미제출",
     "GitHub HW19 폴더에 날씨 데이터 분석 파일만 있어 hw19 제출 없음",
     "GitHub미제출", False, False, False, False),

    ("노영호",       "202515422", "nyh060208-droid",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("문치웅",       "202446401", "moonchiwoong",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("박정호",       "202525236", "amumu-master",
     "-",  "이슈",
     "[hw19_3.py] f-string 중첩으로 SyntaxError 발생",
     "이슈", True, False, False, False),

    ("박혜리",       "202515428", "hyeri6213-ui",
     "-",  "이슈",
     "[hangman.py] 게임 시작 시 print(letters)로 정답 단어 그대로 노출",
     "이슈", True, False, False, False),

    ("배예진",       "202515430", "yejijinn",
     "-",  "이슈복수",
     "[hangman.py] trial -= 1이 조건 밖에서 항상 실행 — 맞혀도 trial 감소; [lotto.py] print가 for 루프 밖에 있어 마지막 회차만 출력됨",
     "이슈복수", True, False, True, False),

    ("서연우",       "202310070", "gogumahotteok",
     "-",  "경미한 이슈",
     '[hangman_game.py] "apple"과 "focus" 사이 쉼표 누락으로 "applefocus" 문자열 결합 버그 (minor)',
     "경미한 이슈", False, False, False, False),

    ("성제현",       "202525237", "jehyunseong0414",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("신준서",       "202310076", "shinjunseo",
     "-",  "이슈",
     "[lotto.py] range(1,45)로 45 제외 — 1~44 범위만 추출됨",
     "이슈", False, False, True, False),

    ("여주엽",       "202423684", "duwnduq0324",
     "-",  "이슈",
     "[19-4.py] random.randint 반복으로 로또 중복 번호 허용",
     "이슈", False, False, True, False),

    ("오태헌",       "202310081", "taeh2004",
     "-",  "경미한 이슈",
     "이상없음 — countdown 사용자 입력 없이 5 고정 (minor)",
     "경미한 이슈", False, False, False, False),

    ("옥상훈",       "202525243", "sangsangbar34",
     "-",  "이슈복수",
     "[hangman.py] f-string 중첩으로 SyntaxError; [lotto_num.py] 6개 각각 randint 추출로 중복 번호 허용",
     "이슈복수", True, False, True, False),

    ("윤건영",       "202210090", "yunkunyoung2003",
     "미제출", "미제출",
     "—",
     "미제출", False, False, False, False),

    ("이강변",       "202410103", "LeeKangByeon",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("이정제",       "202220633", "cbsksbx-png",
     "-",  "이슈복수",
     "[task19_03.py] 단어 순서대로 체크하는 방식으로 행맨 규칙 미준수; [task19_04.py] 중복 방지 로직 불완전 — 중복 가능",
     "이슈복수", True, False, True, False),

    ("이준혁",       "202420965", "jhyeok2003",
     "-",  "경미한 이슈",
     "이상없음 — 각 프로그램 사용자 입력 없이 고정값 사용 (minor)",
     "경미한 이슈", False, False, False, False),

    ("이환",         "202515448", "202515448-maker",
     "-",  "이슈",
     "[hw19_3.py] PySimpleGUI 외부 라이브러리 및 word.csv 필요 — 환경 미충족 시 실행 불가",
     "이슈", True, False, False, False),

    ("정가경",       "202515452", "jkk728",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("정보성",       "202446402", "envns123",
     "-",  "이슈",
     "[homework19_3.py] hangman() 함수 내부 로직 없음 — main()은 pass로 미완성",
     "이슈", True, False, False, False),

    ("조인현",       "202515455", "Hyeoniverse129",
     "미제출", "이슈",
     "[lec13_homework04~07.py] 4개 파일 모두 SyntaxError (콜론 누락, 문자열 미완성 등)",
     "LMS 미제출", True, False, False, False),

    ("조혜윤",       "202515457", "hyeyun0201",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("지형기",       "202310104", "IMJHK0720",
     "-",  "이상없음",
     "이상없음",
     "이상없음", False, False, False, False),

    ("최보성",       "202525257", "choibosung14-dot",
     "미제출", "미제출",
     "—",
     "미제출", False, False, False, False),

    ("잔바예프 다니야르", "202595558", "(없음)",
     "미제출", "이상없음",
     "이상없음",
     "LMS 미제출", False, False, False, False),
]

# ── 점수 계산 ────────────────────────────────────────────────────
def calc_score(s):
    """(s1_hangman, s2_gugudan, s3_lotto, s4_alt, 우수, 최종) 반환"""
    종합 = s[6]
    base = 0 if 종합 == "미제출" else 5
    s1 = -0.5 if s[7]  else None   # 행맨 이슈
    s2 = -0.5 if s[8]  else None   # 구구단 이슈
    s3 = -0.5 if s[9]  else None   # 로또 이슈
    s4 = -0.5 if s[10] else None   # 행맨 대체구현
    우수 = None  # 수동 기입
    deduct = sum(v for v in [s1, s2, s3, s4] if v is not None)
    final = base + deduct if base > 0 else 0
    # 소수점 정리
    final = round(final, 1)
    return s1, s2, s3, s4, 우수, final if base > 0 or final != 0 else 0


# ── Excel 빌드 ────────────────────────────────────────────────────
def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "hw19 점검결과"

    LAST = "T"   # 총 20열

    # ── 행1: 제목 ──────────────────────────────────────────────────
    ws.merge_cells(f"A1:{LAST}1")
    wcell(ws, 1, 1,
          "📋 프원실(2026) 과제19 점검 리포트 (전체)",
          bold=True, sz=11)
    ws.row_dimensions[1].height = 20

    # ── 행2: 부제목 ───────────────────────────────────────────────
    ws.merge_cells(f"A2:{LAST}2")
    wcell(ws, 2, 1,
          f"2026년 1학기  |  과제19: 카운트다운 · 구구단 문제집 · 행맨 · 로또번호추출기 (4개 이상 완성)  |  "
          f"제출 5 | 오류 -0.5 | 우수 1  |  점검일: {TODAY}",
          sz=9, fg="555555")
    ws.row_dimensions[2].height = 15

    # ── 행3: 그룹 헤더 ────────────────────────────────────────────
    # A~C: 기본 정보
    ws.merge_cells("A3:C3")
    wcell(ws, 3, 1, "기본 정보", bold=True, sz=9)
    # D~H: 제출 현황
    ws.merge_cells("D3:H3")
    wcell(ws, 3, 4, "제출 현황", bold=True, sz=9)
    # I~M: 코드 체크
    ws.merge_cells("I3:M3")
    wcell(ws, 3, 9, "코드 체크", bold=True, sz=9)
    # N~T: 점수
    ws.merge_cells("N3:T3")
    wcell(ws, 3, 14, "제출 5 | 오류 -0.5 | 우수 1", bold=True, sz=9)
    ws.row_dimensions[3].height = 18

    # ── 행4: 열 헤더 ─────────────────────────────────────────────
    headers = [
        "이름", "학번", "GitHub ID",          # A B C
        "LMS\n제출", "GitHub\n제출",           # D E
        "로직 이슈 (GitHub)",                  # F
        "최종\n종합",                           # G
        "",                                    # H (빈칸)
        "①행맨\n이슈",                         # I
        "②구구단\n이슈",                        # J
        "③로또\n이슈",                          # K
        "④행맨\n대체구현",                      # L
        "",                                    # M (빈칸)
        "①행맨\n오류",                         # N
        "②구구단\n오류",                        # O
        "③로또\n오류",                          # P
        "④대체\n구현",                          # Q
        "",                                    # R (빈칸)
        "우수\n학생",                           # S
        "최종",                                # T
    ]
    for col, h in enumerate(headers, 1):
        wcell(ws, 4, col, h, bold=True, sz=9)
    ws.row_dimensions[4].height = 40

    # ── 행5~: 학생 데이터 ────────────────────────────────────────
    for row_idx, s in enumerate(STUDENTS, 5):
        이름, 학번, 깃헙아이디 = s[0], s[1], s[2]
        lms, github, 이슈, 종합 = s[3], s[4], s[5], s[6]
        h_행맨, i_구구단, j_로또, k_대체 = s[7], s[8], s[9], s[10]

        is_missing = 종합 == "미제출"
        row_bg = GRAY if is_missing else None

        # A B C
        for col, val in enumerate([이름, 학번, 깃헙아이디], 1):
            wcell(ws, row_idx, col, val, bg=row_bg, sz=9)

        # D: LMS
        lms_bg = GRAY if lms == "미제출" else None
        wcell(ws, row_idx, 4, lms, bg=lms_bg, sz=9)

        # E: GitHub 제출
        gh_bg = GRAY if github == "미제출" else None
        wcell(ws, row_idx, 5, github, bg=gh_bg, sz=9)

        # F: 로직이슈
        wcell(ws, row_idx, 6, 이슈, h="left", sz=9)

        # G: 최종종합
        g_bg = GRAY if is_missing else None
        wcell(ws, row_idx, 7, 종합, bold=True, bg=g_bg, sz=9)

        # H: 빈칸
        ws.cell(row_idx, 8).border = _border()

        # I J K L: 코드체크
        for col_offset, checked in enumerate([h_행맨, i_구구단, j_로또, k_대체]):
            col = 9 + col_offset
            wcell(ws, row_idx, col, "✅" if checked else "—", sz=9)

        # M: 빈칸
        ws.cell(row_idx, 13).border = _border()

        # N O P Q: 점수 감점
        s1, s2, s3, s4, 우수, final = calc_score(s)
        for col_offset, val in enumerate([s1, s2, s3, s4]):
            col = 14 + col_offset
            wcell(ws, row_idx, col, val if val is not None else None, sz=9)

        # R: 빈칸
        ws.cell(row_idx, 18).border = _border()

        # S: 우수
        wcell(ws, row_idx, 19, 우수, sz=9)

        # T: 최종점수
        wcell(ws, row_idx, 20, final if not is_missing else 0, bold=True, sz=9)

        ws.row_dimensions[row_idx].height = 60

    # ── 열 너비 ──────────────────────────────────────────────────
    col_widths = {
        "A": 10, "B": 12, "C": 16,
        "D": 6,  "E": 12, "F": 50, "G": 12,
        "H": 2,
        "I": 10, "J": 10, "K": 10, "L": 10,
        "M": 2,
        "N": 8,  "O": 8,  "P": 8,  "Q": 8,
        "R": 2,
        "S": 6,  "T": 6,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A5"

    wb.save(OUT_PATH)
    print(f"✅ 저장 완료: {OUT_PATH}")


if __name__ == "__main__":
    build()

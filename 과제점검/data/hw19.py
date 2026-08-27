"""
hw19.py — 과제19 데이터
과제 요구사항: 카운트다운 / 구구단 문제집 / 행맨 / 로또번호추출기 4개 이상 완성
"""

CHECK_COLS = [
    "행맨\n이슈",
    "구구단\n이슈",
    "로또\n이슈",
    "행맨\n대체구현",
]

STUDENTS = [
    # (이름, 학번, GitHub_ID, LMS, GitHub, 로직이슈, H, I, J, K)
    ("강태현", "202515406", "kth0208",
     "✅", "🟡 경고",
     "[gugudan1.py] 점수 출력이 for 루프 안에서 문제마다 출력됨",
     False, True, False, False),

    ("김근영", "202515408", "zbr2s5rjc9-sys",
     "✅", "🟡 경고",
     "[countdown.py] 사용자 입력 없이 5부터 고정 카운트 (minor); [lotto.py] 로또 로직이 main() 밖에서 실행됨 (minor)",
     False, False, False, False),

    ("김도영", "202210033", "Doyoung03",
     "—", "🟠 실행 오류",
     "[hw3.py] f-string 중첩으로 SyntaxError 발생 / 맞혀도 trial 감소하는 행맨 로직 오류",
     True, False, False, False),

    ("김민성", "202321610", "Kim-cloud-ai",
     "✅", "🔴 이슈 복수",
     "[homework19-2.py] count+=0으로 틀린 횟수 집계 안 됨; [homework19-3.py] 행맨 대신 약어 맞추기 구현 (행맨 미구현); [homework19-4.py] 로또 중복 번호 허용",
     False, True, True, True),

    ("김선우", "202515414", "SeonU777",
     "✅", "🟡 경고",
     "[hang_man.py] w_list[count] 순서 비교로 행맨 규칙 미준수; [count_down.py] end=\"\" 로 숫자 이어붙여 출력 (minor)",
     True, False, False, False),

    ("김선호", "202515415", "fnfnf2145",
     "—", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("김예지", "202246491", "dpwl829",
     "—", "⛔ 미제출",
     "—",
     False, False, False, False),

    ("김준혁", "202310054", "kimjunhyeok04",
     "✅", "🟡 폴더 미정리",
     "[HW19] .py 없이 ZIP 파일로만 제출 (minor)",
     False, False, False, False),

    ("김현준", "202310055", "202310055HJ",
     "✅", "⛔ 미제출",
     "GitHub HW19 폴더에 날씨 데이터 분석 파일만 있어 hw19 제출 없음",
     False, False, False, False),

    ("노영호", "202515422", "nyh060208-droid",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("문치웅", "202446401", "moonchiwoong",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("박정호", "202525236", "amumu-master",
     "❌", "🟠 실행 오류",
     "[hw19_3.py] f-string 중첩으로 SyntaxError 발생",
     True, False, False, False),

    ("박혜리", "202515428", "hyeri6213-ui",
     "✅", "🟡 경고",
     "[hangman.py] 게임 시작 시 print(letters)로 정답 단어 그대로 노출",
     True, False, False, False),

    ("배예진", "202515430", "yejijinn",
     "✅", "🔴 이슈 복수",
     "[hangman.py] trial -= 1이 조건 밖에서 항상 실행 — 맞혀도 trial 감소; [lotto.py] print가 for 루프 밖에 있어 마지막 회차만 출력됨",
     True, False, True, False),

    ("서연우", "202310070", "gogumahotteok",
     "✅", "🟡 경고",
     "[hangman_game.py] \"apple\"과 \"focus\" 사이 쉼표 누락으로 \"applefocus\" 문자열 결합 버그 (minor)",
     True, False, False, False),

    ("성제현", "202525237", "jehyunseong0414",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("신준서", "202310076", "shinjunseo",
     "✅", "🟡 경고",
     "[lotto.py] range(1,45)로 45 제외 — 1~44 범위만 추출됨",
     False, False, True, False),

    ("여주엽", "202423684", "duwnduq0324",
     "✅", "🟡 경고",
     "[19-4.py] random.randint 반복으로 로또 중복 번호 허용",
     False, False, True, False),

    ("오태헌", "202310081", "taeh2004",
     "✅", "🟢 이상 없음",
     "이상 없음 — countdown 사용자 입력 없이 5 고정 (minor)",
     False, False, False, False),

    ("옥상훈", "202525243", "sangsangbar34",
     "❌", "🟠 실행 오류",
     "[hangman.py] f-string 중첩으로 SyntaxError; [lotto_num.py] 6개 각각 randint 추출로 중복 번호 허용",
     True, False, True, False),

    ("윤건영", "202210090", "yunkunyoung2003",
     "—", "⛔ 미제출",
     "—",
     False, False, False, False),

    ("이강변", "202410103", "LeeKangByeon",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("이정제", "202220633", "cbsksbx-png",
     "✅", "🔴 이슈 복수",
     "[task19_03.py] 단어 순서대로 체크하는 방식으로 행맨 규칙 미준수; [task19_04.py] 중복 방지 로직 불완전 — 중복 가능",
     True, False, True, False),

    ("이준혁", "202420965", "jhyeok2003",
     "✅", "🟢 이상 없음",
     "이상 없음 — 각 프로그램 사용자 입력 없이 고정값 사용 (minor)",
     False, False, False, False),

    ("이환", "202515448", "202515448-maker",
     "✅", "🟡 경고",
     "[hw19_3.py] PySimpleGUI 외부 라이브러리 및 word.csv 필요 — 환경 미충족 시 실행 불가",
     True, False, False, False),

    ("정가경", "202515452", "jkk728",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("정보성", "202446402", "envns123",
     "✅", "🟡 경고",
     "[homework19_3.py] hangman() 함수 내부 로직 없음 — main()은 pass로 미완성",
     True, False, False, False),

    ("조인현", "202515455", "Hyeoniverse129",
     "—", "🟠 실행 오류",
     "[lec13_homework04~07.py] 4개 파일 모두 SyntaxError (콜론 누락, 문자열 미완성 등)",
     True, False, False, False),

    ("조혜윤", "202515457", "hyeyun0201",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("지형기", "202310104", "IMJHK0720",
     "✅", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),

    ("최보성", "202525257", "choibosung14-dot",
     "—", "⛔ 미제출",
     "—",
     False, False, False, False),

    ("잔바예프 다니야르", "202595558", "(없음)",
     "—", "🟢 이상 없음",
     "이상 없음",
     False, False, False, False),
]

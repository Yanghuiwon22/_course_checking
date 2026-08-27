import pandas as pd

file_path = "출석부_프원실(2026).xlsx"
import pandas as pd

# 학생 이름 리스트 직접 입력
all_students = [
    '김도영', '윤건영', '이정제', '김예지', '김준혁', '김현준',
    '서연우', '신준서', '오태헌', '지형기', '김민성', '이강변',
    '이준혁', '여주엽', '문치웅', '정보성', '강태현', '김근영',
    '김선우', '김선호', '노영호', '박혜리', '배예진', '이환',
    '정가경', '조인현', '조혜윤', '박정호', '성제현', '옥상훈',
    '최보성', '잔바예프 다니야르'
]

name_map = {
    'Janbayev': '잔바예프 다니야르',
}

xl = pd.ExcelFile(file_path)

# 탭별 이름 추출
tab_names = {}

for sheet in xl.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    block1 = df.iloc[1:7, 0:7]  # A2:G7 (기본 참석)
    block2 = df.iloc[10:30, 1:2]  # A10:G15 (추가 참석) ← 범위 알려주세요

    combined = pd.concat([block1, block2])
    col = combined.values.flatten()

    col = [name_map.get(x, x) for x in col]

    tab_names[sheet] = [x for x in col if pd.notna(x)]

# 결과 데이터프레임 생성
result = pd.DataFrame(index=all_students, columns=xl.sheet_names)
result.index.name = '이름'

for sheet, names in tab_names.items():
    result[sheet] = result.index.map(lambda x: '참석' if x in names else '미참석')

print(result)

# 한 번도 참석 안 한 학생 (모든 탭에서 미참석)
never_attended = [s for s in all_students if all(s not in names for names in tab_names.values())]
if never_attended:
    print("\n⚠️ 어느 탭에도 없는 학생:")
    for name in never_attended:
        print(f"  - {name}")

# 엑셀로 저장
result.to_csv("출석결과.csv")
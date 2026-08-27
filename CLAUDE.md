# 프로그래밍원리와실습 과제 점검 프로젝트

이 프로젝트는 학생들이 과제를 GitHub 저장소에 업로드하면, 매주 일요일 오후 12시에 pull 하여 과제 제출물의 업데이트를 체크하는 작업을 목표로 한다.

> 상세 작업 가이드(리포트 형식, 판정 기준, 수강생 목록 등)는 [`과제점검/CLAUDE.md`](과제점검/CLAUDE.md)를 참조한다.

---

## 폴더 구조

```
프원실/
├── CLAUDE.md                           ← 이 파일 (프로젝트 개요)
├── requirements.txt
├── lms 제출물/                          ← LMS에서 다운받은 제출물 (학생별 폴더)
├── github 과제점검/                     ← GitHub에서 clone한 학생 저장소 모음
│   └── {학생이름 (영문)}/               ← 학생별 저장소 (clone_from_excel.py로 생성)
│       └── hw07/ hw08/ ...              ← 학생이 업로드한 과제 폴더
├── 근무일지/
└── 과제점검/                            ← 점검 스크립트 및 결과물
    ├── CLAUDE.md                        ← ✅ 상세 작업 가이드 (메인 레퍼런스)
    ├── credentials.json                 ← Google Sheets 인증 키
    ├── .env                             ← 환경 변수 (LMS 계정 등)
    ├── 프원실(2026)_과제확인.xlsx        ← 수강생 명단 + GitHub clone URL
    │                                       (D열: 학번, E열: 이름, F열: GitHub ID, M열: clone 명령어)
    ├── config.py                        ← 공통 설정 (경로, 상수 등)
    ├── utils.py                         ← 공통 유틸리티 함수
    ├── auto_login.py                    ← LMS 로그인 & 제출물 다운로드
    ├── clone_from_excel.py              ← GitHub 저장소 clone/pull
    ├── check_submissions.py             ← LMS 제출 여부 확인 → Google Sheets 업데이트
    ├── auto_check_github.py             ← GitHub 과제 점검 → xlsx 리포트 생성
    │                                       (run() 함수로 임포트 가능 / CLI 직접 실행도 가능)
    ├── build_report.py                  ← 최종 통합 리포트 생성 (공통 모듈)
    ├── weekly_runner.py                 ← 주차별 자동화 실행 스크립트
    ├── 과제모음/                        ← 과제 내용 텍스트 파일
    │   ├── TEMPLATE.txt                 ← 과제 텍스트 작성 양식
    │   └── hw02.txt ~ hw21.txt          ← 과제별 문제 설명
    ├── output/                          ← 점검 중간 결과물 (주차별 하위폴더)
    ├── data/                            ← 데이터 모듈 (hw_template.py 등)
    └── 과제XX_점검리포트.xlsx           ← 생성된 점검 리포트 (과제별)
```

---

## 스크립트 역할 및 실행 방법

> 모든 스크립트는 `과제점검/` 폴더에서 실행한다.

### `weekly_runner.py` — 주차별 자동화 실행 (권장)

```bash
cd 과제점검
python weekly_runner.py
```

### 개별 스크립트 (단계별 수동 실행)

#### `auto_login.py` — LMS 로그인 & 제출물 다운로드
LMS에 자동 로그인하여 해당 과제 제출물을 `lms 제출물/` 폴더에 다운로드한다.

#### `clone_from_excel.py` — GitHub 저장소 clone/pull
`프원실(2026)_과제확인.xlsx`의 M열(clone 명령어)을 기준으로 학생 저장소를 `github 과제점검/` 폴더에 clone하거나, 이미 존재하면 pull한다.

```bash
python clone_from_excel.py
python clone_from_excel.py --dry-run  # 실제 실행 없이 확인만
```

#### `check_submissions.py` — LMS 제출 여부 확인
`lms 제출물/` 폴더를 탐색해 학생별 hw 제출 현황을 Google Sheets에 업데이트한다.
- 미제출: 빨간색 / 지각 제출: 노란색

```bash
python check_submissions.py
```

#### `auto_check_github.py` — GitHub 과제 자동 점검
특정 과제 번호를 기준으로 학생별 GitHub 저장소를 탐색·분석하여 xlsx 리포트를 생성한다.

```bash
python auto_check_github.py --hw 7
python auto_check_github.py --hw 7 --out ../과제07_점검리포트_github.xlsx
```

#### `build_report.py` — 최종 통합 리포트 생성
LMS 점검 결과와 GitHub 점검 결과를 합쳐 최종 xlsx 리포트를 생성한다.

---

## 과제 점검 절차

사용자가 특정 과제 번호 점검을 요청하면 아래 순서로 진행한다.

1. **과제 내용 확인** — `과제점검/과제모음/hw{번호}.txt` 파일로 과제 요구사항을 파악한다.
2. **폴더 구조 점검** — `github 과제점검/{학생명}/` 하위에서 해당 과제 번호 폴더를 탐색한다.
3. **Python 코드 점검** — 구문 검사(ast.parse) 및 실행 테스트를 수행한다.
4. **리포트 생성** — 학생별 이슈를 정리하여 xlsx 리포트로 저장한다.

### 과제 폴더 탐색 기준
폴더 형식은 엄격히 정해져 있지 않으며, 아래 기준으로 판정한다.
**폴더명이 비표준이더라도 과제 번호가 포함되어 있으면 정상으로 처리한다.**

| 판정 | 조건 |
|------|------|
| ✅ 정상 | `hw07`, `homework07`, `과제07_...`, `task_07`, `assignment12` 등 과제 번호가 포함된 모든 폴더명 |
| ❌ 미정리 | 폴더 없이 루트에 `.py` 파일 직접 제출 |
| — 미제출 | 해당 과제 파일 없음 |

### 점검 항목
- 과제 번호 폴더별로 정리되어 있지 않은 학생 목록 우선 정리
- Python 코드 구문 검사 (ast.parse) 및 실행 테스트 (제한 시간 3초)
- 과제 내용 기준으로 로직 이슈 여부 확인 (수동 기입 — xlsx F열)
- 학생별 이슈 사항 리포트 (폴더 미정리 / 코드 오류 / 미제출 등)

---

## LMS 채점 기준

| 값 | 조건 |
|----|------|
| `✅` | `.py` 파일이 있고 `ast.parse()` 구문 검사 통과, 또는 `onlinetext.html`만 존재 (텍스트 제출로 간주) |
| `❌` | `.py` 파일은 있으나 `SyntaxError` 발생 |
| `—` | 아예 미제출 |

---

## 수강생 명단 파일

| 파일 | 위치 | 용도 |
|------|------|------|
| `프원실(2026)_과제확인.xlsx` | `과제점검/` | 학번(D열), 이름(E열), GitHub ID(F열), clone URL(M열) — 스크립트 기준 파일 |

"""
lms_제출_점검 Google Sheets 제출/미제출 현황 업데이트 스크립트.

사용법:
    cd 과제점검
    python check_submissions.py
"""

from pathlib import Path
import openpyxl
import gspread
from gspread.utils import rowcol_to_a1

from config import (SPREADSHEET_ID, SHEET_GID, LMS_DIR as STUDENT_DIR,
                    ROSTER_PATH, CREDS_PATH, COL_STUDENT_NAME as ROSTER_COL)
from utils import normalize_name

def hex_rgb(h):
    h = h.lstrip('#')
    return {'red': int(h[0:2], 16) / 255, 'green': int(h[2:4], 16) / 255, 'blue': int(h[4:6], 16) / 255}


def cell_range(gid, row, col):
    return {
        'sheetId': gid,
        'startRowIndex': row - 1, 'endRowIndex': row,
        'startColumnIndex': col - 1, 'endColumnIndex': col,
    }


def row_range(gid, row, start_col, end_col):
    return {
        'sheetId': gid,
        'startRowIndex': row - 1, 'endRowIndex': row,
        'startColumnIndex': start_col - 1, 'endColumnIndex': end_col,
    }


def iter_col_ranges(cols):
    """1-based column indices -> contiguous [start, end_exclusive] ranges."""
    if not cols:
        return []
    cols = sorted(set(cols))
    ranges = []
    start = prev = cols[0]
    for c in cols[1:]:
        if c == prev + 1:
            prev = c
            continue
        ranges.append((start, prev + 1))
        start = prev = c
    ranges.append((start, prev + 1))
    return ranges


def apply_formatting(sh, sheet_gid, num_data_rows, num_data_cols,
                     red_cells, plain_cells, yellow_cells, new_rows, new_cols, header_created):
    """신규 행/열만 서식을 적용합니다."""
    if not new_rows and not new_cols and not header_created:
        return
    gid = sheet_gid
    R = []  # requests

    # ── 시트 속성 (탭 색, 틀 고정) ──
    R.append({'updateSheetProperties': {
        'properties': {
            'sheetId': gid,
            'tabColorStyle': {'rgbColor': hex_rgb('2C3E50')},
            'gridProperties': {'frozenRowCount': 1, 'frozenColumnCount': 1},
        },
        'fields': 'tabColorStyle,gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    }})

    # ── 행 높이 ──
    R.append({'updateDimensionProperties': {
        'range': {'sheetId': gid, 'dimension': 'ROWS', 'startIndex': 0, 'endIndex': 1},
        'properties': {'pixelSize': 30}, 'fields': 'pixelSize',
    }})
    for row in new_rows:
        R.append({'updateDimensionProperties': {
            'range': {'sheetId': gid, 'dimension': 'ROWS', 'startIndex': row - 1, 'endIndex': row},
            'properties': {'pixelSize': 25}, 'fields': 'pixelSize',
        }})

    # ── 열 너비 ──
    R.append({'updateDimensionProperties': {
        'range': {'sheetId': gid, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1},
        'properties': {'pixelSize': 150}, 'fields': 'pixelSize',
    }})
    for start_col, end_col in iter_col_ranges(new_cols):
        R.append({'updateDimensionProperties': {
            'range': {'sheetId': gid, 'dimension': 'COLUMNS',
                      'startIndex': start_col - 1, 'endIndex': end_col - 1},
            'properties': {'pixelSize': 80}, 'fields': 'pixelSize',
        }})

    # ── 헤더 행 서식 (행 1) ──
    header_cols = set(new_cols)
    if header_created:
        header_cols.add(1)
    for start_col, end_col in iter_col_ranges(header_cols):
        R.append({'repeatCell': {
            'range': {'sheetId': gid, 'startRowIndex': 0, 'endRowIndex': 1,
                      'startColumnIndex': start_col - 1, 'endColumnIndex': end_col - 1},
            'cell': {'userEnteredFormat': {
                'backgroundColor': hex_rgb('2C3E50'),
                'horizontalAlignment': 'CENTER',
                'verticalAlignment': 'MIDDLE',
                'textFormat': {
                    'foregroundColor': hex_rgb('FFFFFF'),
                    'fontFamily': 'Arial', 'fontSize': 11, 'bold': True,
                },
                'borders': {k: {'style': 'SOLID_MEDIUM', 'colorStyle': {'rgbColor': hex_rgb('1A252F')}}
                            for k in ('top', 'bottom', 'left', 'right')},
            }},
            'fields': 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,borders)',
        }})

    # ── A열 (이름) 서식 ──
    for row in new_rows:
        R.append({'repeatCell': {
            'range': {'sheetId': gid, 'startRowIndex': row - 1, 'endRowIndex': row,
                      'startColumnIndex': 0, 'endColumnIndex': 1},
            'cell': {'userEnteredFormat': {
                'backgroundColor': hex_rgb('ECF0F1'),
                'horizontalAlignment': 'LEFT',
                'verticalAlignment': 'MIDDLE',
                'textFormat': {
                    'foregroundColor': hex_rgb('2C3E50'),
                    'fontFamily': 'Arial', 'fontSize': 10, 'bold': True,
                },
            }},
            'fields': 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)',
        }})

    # ── 데이터 셀 기본 서식 (신규 열 전체, 신규 행 전체) ──
    for start_col, end_col in iter_col_ranges(new_cols):
        R.append({'repeatCell': {
            'range': {'sheetId': gid, 'startRowIndex': 1, 'endRowIndex': 1 + num_data_rows,
                      'startColumnIndex': start_col - 1, 'endColumnIndex': end_col - 1},
            'cell': {'userEnteredFormat': {
                'horizontalAlignment': 'CENTER',
                'verticalAlignment': 'MIDDLE',
                'textFormat': {
                    'foregroundColor': hex_rgb('6C757D'),
                    'fontFamily': 'Arial', 'fontSize': 10, 'bold': False,
                },
                'borders': {k: {'style': 'SOLID', 'colorStyle': {'rgbColor': hex_rgb('DEE2E6')}}
                            for k in ('top', 'bottom', 'left', 'right')},
            }},
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,borders)',
        }})
    for row in new_rows:
        R.append({'repeatCell': {
            'range': {'sheetId': gid, 'startRowIndex': row - 1, 'endRowIndex': row,
                      'startColumnIndex': 1, 'endColumnIndex': 1 + num_data_cols},
            'cell': {'userEnteredFormat': {
                'horizontalAlignment': 'CENTER',
                'verticalAlignment': 'MIDDLE',
                'textFormat': {
                    'foregroundColor': hex_rgb('6C757D'),
                    'fontFamily': 'Arial', 'fontSize': 10, 'bold': False,
                },
                'borders': {k: {'style': 'SOLID', 'colorStyle': {'rgbColor': hex_rgb('DEE2E6')}}
                            for k in ('top', 'bottom', 'left', 'right')},
            }},
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,borders)',
        }})

    # ── 홀짝 행 배경색 (신규 열 전체 / 신규 행 전체) ──
    for i in range(num_data_rows):
        row = i + 2  # 시트 1-indexed, 1행=헤더
        bg = 'FFFFFF' if i % 2 == 0 else 'F7F9FA'
        for start_col, end_col in iter_col_ranges(new_cols):
            R.append({'repeatCell': {
                'range': {'sheetId': gid, 'startRowIndex': row - 1, 'endRowIndex': row,
                          'startColumnIndex': start_col - 1, 'endColumnIndex': end_col - 1},
                'cell': {'userEnteredFormat': {'backgroundColor': hex_rgb(bg)}},
                'fields': 'userEnteredFormat.backgroundColor',
            }})
    for row in new_rows:
        bg = 'FFFFFF' if (row - 2) % 2 == 0 else 'F7F9FA'
        R.append({'repeatCell': {
            'range': {'sheetId': gid, 'startRowIndex': row - 1, 'endRowIndex': row,
                      'startColumnIndex': 1, 'endColumnIndex': 1 + num_data_cols},
            'cell': {'userEnteredFormat': {'backgroundColor': hex_rgb(bg)}},
            'fields': 'userEnteredFormat.backgroundColor',
        }})

    # ── 셀별 서식 (미제출/지각/정상) ──
    for row, col in red_cells:
        R.append({'repeatCell': {
            'range': cell_range(gid, row, col),
            'cell': {'userEnteredFormat': {
                'backgroundColor': hex_rgb('F8D7DA'),
                'textFormat': {
                    'foregroundColor': hex_rgb('842029'),
                    'fontFamily': 'Arial', 'fontSize': 10, 'bold': True,
                },
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)',
        }})

    for row, col in yellow_cells:
        R.append({'repeatCell': {
            'range': cell_range(gid, row, col),
            'cell': {'userEnteredFormat': {
                'backgroundColor': hex_rgb('FFF3CD'),
                'textFormat': {
                    'foregroundColor': hex_rgb('856404'),
                    'fontFamily': 'Arial', 'fontSize': 10, 'bold': True,
                },
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat)',
        }})

    for row, col in plain_cells:
        R.append({'repeatCell': {
            'range': cell_range(gid, row, col),
            'cell': {'userEnteredFormat': {
                'textFormat': {
                    'foregroundColor': hex_rgb('6C757D'),
                    'fontFamily': 'Arial', 'fontSize': 10, 'bold': False,
                },
            }},
            'fields': 'userEnteredFormat.textFormat',
        }})

    sh.batch_update({'requests': R})


def update_submission_checklist(spreadsheet_id, sheet_gid, student_names, student_dir_root, creds_path):
    gc = gspread.service_account(filename=str(creds_path))
    sh = gc.open_by_key(spreadsheet_id)
    ws = next(w for w in sh.worksheets() if w.id == sheet_gid)

    all_values = ws.get_all_values()

    if all_values and all_values[0]:
        header = all_values[0]
        name_to_row = {}
        for i, row in enumerate(all_values[1:], start=2):
            if row and row[0]:
                name_to_row[row[0]] = i
        next_col = len(header) + 1
        current_max_row = len(all_values)
        header_created = False
    else:
        header = []
        name_to_row = {}
        next_col = 2
        current_max_row = 1
        header_created = True

    # 헤더에서 hw 열 위치 파악 (대소문자/공백 무시: 'HW 02' → 'hw02')
    hw_col: dict[str, int] = {}
    for col_idx, h in enumerate(header, start=1):
        norm = str(h).lower().replace(' ', '')
        if norm.startswith('hw'):
            hw_col[norm] = col_idx

    # hw 목록 수집
    hw_set: set[str] = set()
    for sd in student_dir_root.iterdir():
        if sd.is_dir():
            for sub in sd.iterdir():
                if sub.is_dir() and sub.name.startswith('hw'):
                    hw_set.add(sub.name)
    all_hw = sorted(hw_set)

    new_hw = [hw for hw in all_hw if hw not in hw_col]
    for i, hw in enumerate(new_hw):
        hw_col[hw] = next_col + i

    value_updates = []
    red_cells = []
    plain_cells = []
    yellow_cells = []

    if not header:
        value_updates.append({'range': 'A1', 'values': [['학생']]})

    for i, hw in enumerate(new_hw):
        value_updates.append({'range': rowcol_to_a1(1, next_col + i), 'values': [[hw]]})

    new_rows = []
    for name in student_names:
        name = name.split(' ')[0]
        is_new = name not in name_to_row
        if not is_new:
            row_idx = name_to_row[name]
            hws_to_fill = new_hw
        else:
            current_max_row += 1
            row_idx = current_max_row
            name_to_row[name] = row_idx
            value_updates.append({'range': rowcol_to_a1(row_idx, 1), 'values': [[name]]})
            new_rows.append(row_idx)
            hws_to_fill = all_hw

        for hw in hws_to_fill:
            hw_dir = student_dir_root / name / hw
            has_files = hw_dir.is_dir() and any(f.is_file() for f in hw_dir.iterdir())
            col = hw_col[hw]
            if has_files:
                value_updates.append({'range': rowcol_to_a1(row_idx, col), 'values': [['-']]})
                plain_cells.append((row_idx, col))
            else:
                value_updates.append({'range': rowcol_to_a1(row_idx, col), 'values': [['미제출']]})
                red_cells.append((row_idx, col))

        # 기존 열: 이전에 미제출 -> 현재 제출이면 "지각"으로 표기
        if not is_new:
            row_values = all_values[row_idx - 1] if row_idx - 1 < len(all_values) else []
            for hw in all_hw:
                if hw in new_hw:
                    continue
                col = hw_col[hw]
                prev_val = row_values[col - 1] if col - 1 < len(row_values) else ''
                if prev_val != '미제출':
                    continue
                hw_dir = student_dir_root / name / hw
                has_files = hw_dir.is_dir() and any(f.is_file() for f in hw_dir.iterdir())
                if has_files:
                    value_updates.append({'range': rowcol_to_a1(row_idx, col), 'values': [['지각']]})
                    yellow_cells.append((row_idx, col))

    if value_updates:
        ws.batch_update(value_updates)

    num_data_rows = current_max_row - 1
    num_data_cols = len(hw_col)

    new_cols = [hw_col[hw] for hw in new_hw]
    if new_rows or new_cols or header_created or yellow_cells:
        print("서식 적용 중...")
        apply_formatting(sh, sheet_gid, num_data_rows, num_data_cols,
                         red_cells, plain_cells, yellow_cells, new_rows, new_cols, header_created)

    msg = "완료"
    if new_hw:
        msg += f" (추가된 hw: {new_hw})"
    print(msg)


def main():
    wb = openpyxl.load_workbook(ROSTER_PATH, read_only=True, data_only=True)
    idx = ROSTER_COL - 1
    names = [normalize_name(row[idx].value) for row in wb.active.iter_rows(min_row=2) if row[idx].value]
    wb.close()

    update_submission_checklist(SPREADSHEET_ID, SHEET_GID, names, STUDENT_DIR, CREDS_PATH)


if __name__ == '__main__':
   main()
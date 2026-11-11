"""매크로 없는 임원 일일 예정사항 양식(.xlsx)을 만듭니다.

한셀처럼 매크로를 쓸 수 없는 곳을 위한 양식입니다. 대시보드 시트의 8칸은 수식으로만 채워지고,
인쇄 설정(A4 · 한 쪽에 맞추기 · 인쇄 영역)은 파일에 들어 있습니다.
TEXTJOIN 같은 새 함수는 쓰지 않습니다(IF·ISNUMBER·FIND·SUBSTITUTE·CHAR·COUNTA·TRIM·CHOOSE·WEEKDAY 등만).

    python3 excel/make_formula_template.py   →  excel/임원일일예정_대시보드_매크로없음.xlsx
"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '임원일일예정_대시보드_매크로없음.xlsx')
FONT = '맑은 고딕'

EXECS = [  # (직위, 성명, 자리 위치, 예시 일정: (시간, 내용, 비고, ※ 참고))
    ('부사장', '이정훈', '윗줄 왼쪽', [
        ('09:00', '인사위원회 @소회의실', '하반기 승진 심사', '16:00 이후 사내 대기'),
        ('10:00', '주간 경영회의 @대회의실', '', ''),
        ('13:30', '감사팀 업무보고', '', ''),
        ('15:00', '외부 감사인 미팅 @접견실', '', '')]),
    ('부사장', '박서연', '윗줄 가운데', [
        ('08:30', '임원 조찬 간담회 @20층 라운지', '', '현장 점검 후 직퇴'),
        ('10:00', '주간 경영회의 @대회의실', '', ''),
        ('14:00~17:00', '영업 현장 점검 (부산지사)', '동행: 영업본부장, 영남지역팀장', '')]),
    ('전무', '최영호', '윗줄 오른쪽', [
        ('10:00', '주간 경영회의 @대회의실', '', ''),
        ('11:00', '신사업 TF 회의', '2027년 신규 진출 시장 검토', ''),
        ('15:00', '협력사 대표 면담 @접견실', '', '')]),
    ('대표이사', '김민수', '가운데 줄 왼쪽 (큰 칸)', [
        ('08:30', '임원 조찬 간담회 @20층 라운지', '참석: 부사장 이하 임원 전원', '11/10(화) 싱가포르 출장 (08:00 인천 출발)'),
        ('10:00', '주간 경영회의 @대회의실\n- 안건: 4분기 실적 전망, 2027년 사업계획 방향', '', ''),
        ('11:30', '신규 고객사 대표 접견 @접견실', '', ''),
        ('12:00', '오찬 (고객사 대표)', '', ''),
        ('14:00', '노사협의회 정기회의 @3층 회의실', '노측 6명, 사측 6명', ''),
        ('16:00', '투자심의위원회 @대회의실', '안건: 제2공장 설비 증설', ''),
        ('17:30', '결재', '', ''),
        ('19:00', '업계 CEO 만찬', '', '')]),
    ('전무', '정다은', '가운데 줄 오른쪽 위', [
        ('09:30', '재무팀 자금 현황 보고', '', '오후 외부 일정 후 복귀'),
        ('10:00', '주간 경영회의 @대회의실', '', ''),
        ('14:00', '주거래 은행 미팅 @외부', '', '')]),
    ('상무', '강태준', '가운데 줄 오른쪽 아래', [
        ('종일', '해외 바이어 상담 (서울사무소)', '베트남 바이어 3개사', '10:00 경영회의 불참 (정다은 전무 대리 보고)')]),
    ('상무', '윤지훈', '아랫줄 왼쪽', [
        ('10:00', '주간 경영회의 @대회의실', '', ''),
        ('13:00', '연구소 과제 점검 @기술연구소', '하반기 중점과제 5건', ''),
        ('16:30', '품질 회의', '', '')]),
    ('상무', '한소영', '아랫줄 오른쪽', [
        ('09:00', '홍보 전략 회의', '', '내일 오전 반차'),
        ('11:00', '언론사 인터뷰 @홍보실', '주제: 친환경 경영', ''),
        ('15:00', '사회공헌 위원회', '', '')]),
]

# 대시보드 칸: (첫 열, 끝 열, 머리 행, 첫 일정 행, 일정 수(칸), ※ 행). 일정 하나 = 두 행(본 줄 + * 비고 줄)
BOXES = [
    (1, 4, 4, 5, 7, 19), (5, 8, 4, 5, 7, 19), (9, 12, 4, 5, 7, 19),
    (1, 8, 21, 22, 13, 48), (9, 12, 21, 22, 6, 34), (9, 12, 35, 36, 6, 48),
    (1, 6, 50, 51, 7, 65), (7, 12, 50, 51, 7, 65),
]
LAST_ROW = 65
REF_ROWS = 10          # ※ 참고 칸에서 모아 보는 줄 수 (입력 2~11행)

THIN = Side(style='thin', color='222222')
DOT = Side(style='dotted', color='888888')
GRAY = PatternFill('solid', fgColor='E7E9EC')


def q(name):
    return "'" + name.replace("'", "''") + "'"


def outline(ws, c1, r1, c2, r2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            b = cell.border
            cell.border = Border(
                left=THIN if c == c1 else b.left, right=THIN if c == c2 else b.right,
                top=THIN if r == r1 else b.top, bottom=THIN if r == r2 else b.bottom)


def put(ws, r, c1, c2, value, size=10, bold=False, align='left', fill=None, shrink=True, color=None):
    cell = ws.cell(r, c1, value)
    cell.font = Font(name=FONT, size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical='center', shrink_to_fit=shrink, wrap_text=False)
    if fill:
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    return cell


def main():
    wb = Workbook()
    dash = wb.active
    dash.title = '대시보드'
    cfg = wb.create_sheet('설정')
    sheet_names = [f'{p} {n}' for p, n, _, _ in EXECS]

    # ---- 설정 ----
    cfg.column_dimensions['A'].width = 16
    cfg.column_dimensions['B'].width = 14
    cfg.column_dimensions['C'].width = 12
    cfg.column_dimensions['D'].width = 24
    cfg.column_dimensions['E'].width = 22
    rows = [('제목', '임원 일일 예정사항'), ('날짜', dt.date.today()), ('왼쪽 위 문구', '')]
    for i, (k, v) in enumerate(rows, start=1):
        cfg.cell(i, 1, k).font = Font(name=FONT, bold=True)
        cell = cfg.cell(i, 2, v)
        cell.font = Font(name=FONT)
        cell.fill = PatternFill('solid', fgColor='FFF9DB')
    cfg['B2'].number_format = 'yyyy-mm-dd'
    for c, h in enumerate(['자리', '직위', '성명', '위치', '입력하는 시트'], start=1):
        cell = cfg.cell(5, c, h)
        cell.font = Font(name=FONT, bold=True)
        cell.fill = GRAY
    for i, (p, n, where, _) in enumerate(EXECS):
        r = 6 + i
        cfg.cell(r, 1, i + 1).font = Font(name=FONT)
        for c, v in ((2, p), (3, n)):
            cell = cfg.cell(r, c, v)
            cell.font = Font(name=FONT)
            cell.fill = PatternFill('solid', fgColor='FFF9DB')
        cfg.cell(r, 4, where).font = Font(name=FONT, color='5D6875')
        cfg.cell(r, 5, sheet_names[i]).font = Font(name=FONT, color='5D6875')
    notes = [
        '사용법',
        '1. 임원마다 시트가 하나씩 있습니다. 한 줄에 일정 하나: 시간 / 내용 / 비고 / ※ 참고.',
        '2. 대시보드 시트는 수식으로 저절로 채워집니다. 대시보드에는 직접 쓰지 마세요.',
        '3. 내용 끝 "@장소" → (장소),  비고 → 다음 줄 작은 글씨 *,  ※ 참고 → 그 칸 맨 밑.',
        '4. 인쇄: 대시보드 시트에서 인쇄하면 A4 한 장에 맞춰 나옵니다 (한 쪽에 맞추기가 들어 있음).',
        '5. 칸마다 담을 수 있는 일정 수: 윗줄·아랫줄 7개, 가운데 큰 칸 13개, 오른쪽 위·아래 6개.',
        '   넘치면 그 칸 맨 밑에 ▲ 표시가 나옵니다.',
        '6. 사람이 바뀌면 위 표의 직위·성명을 고치고, 원하면 시트 이름도 바꾸세요 (수식은 따라갑니다).',
        '7. 두레이 공동편집 파일에서 옮길 때: 같은 사람 시트의 A2:D 범위를 복사해 이 파일의 그 시트 A2에 붙여넣습니다.',
    ]
    for i, t in enumerate(notes):
        cfg.cell(16 + i, 1, t).font = Font(name=FONT, bold=(i == 0))

    # ---- 임원별 입력 시트 ----
    for i, (p, n, _, sched) in enumerate(EXECS):
        ws = wb.create_sheet(sheet_names[i])
        for c, (h, w) in enumerate([('시간', 14), ('내용', 54), ('비고 (작은 글씨 *)', 32), ('※ 참고', 34)], start=1):
            cell = ws.cell(1, c, h)
            cell.font = Font(name=FONT, bold=True)
            cell.fill = GRAY
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(c)].width = w
        for r, row in enumerate(sched, start=2):
            for c, v in enumerate(row, start=1):
                cell = ws.cell(r, c, v if v else None)
                cell.font = Font(name=FONT)
                cell.alignment = Alignment(vertical='top', wrap_text=(c > 1))
        ws.freeze_panes = 'A2'

    # ---- 대시보드 ----
    for c in range(1, 13):
        dash.column_dimensions[get_column_letter(c)].width = 7.9
    heights = {1: 30, 2: 18, 3: 4, 20: 6, 49: 6}
    for (_, _, hr, fr, cap, rr) in BOXES:
        heights.setdefault(hr, 19)
        for k in range(cap):
            heights.setdefault(fr + 2 * k, 15)
            heights.setdefault(fr + 2 * k + 1, 11)
        heights.setdefault(rr, 15)
    for r in range(1, LAST_ROW + 1):
        dash.row_dimensions[r].height = heights.get(r, 15)

    put(dash, 1, 1, 12, '=설정!B1', size=18, bold=True, align='center')
    put(dash, 2, 1, 6, '=IF(설정!B3="","",설정!B3)', size=10)
    put(dash, 2, 7, 12, '=IF(설정!B2="","",YEAR(설정!B2)&". "&MONTH(설정!B2)&". "&DAY(설정!B2)&".("&'
        'CHOOSE(WEEKDAY(설정!B2),"일","월","화","수","목","금","토")&")")', size=10, bold=True, align='right')
    for c in range(1, 13):
        dash.cell(2, c).border = Border(bottom=Side(style='medium', color='000000'))

    for i, (c1, c2, hr, fr, cap, rr) in enumerate(BOXES):
        s = q(sheet_names[i])
        put(dash, hr, c1, c2, f'=TRIM(설정!B{6 + i}&" "&설정!C{6 + i})', size=10.5, bold=True, align='center', fill=GRAY)
        for c in range(c1, c2 + 1):
            dash.cell(hr, c).border = Border(bottom=THIN)
        for k in range(cap):
            src = k + 2          # 입력 시트의 행
            r = fr + 2 * k
            b = f'{s}!B{src}'
            content = (f'IF({b}="","",IF(ISNUMBER(FIND(CHAR(10),{b})),SUBSTITUTE(SUBSTITUTE({b}," @"," · "),CHAR(10)," / "),'
                       f'IF(ISNUMBER(FIND(" @",{b})),SUBSTITUTE({b}," @"," (")&")",{b})))')
            if k == 0:
                content = f'IF(COUNTA({s}!A2:C200)=0,"일정 없음",{content})'
            t = put(dash, r, c1, c1, f'=IF({s}!A{src}="","",{s}!A{src})', size=10, bold=True)
            t.number_format = 'hh:mm'
            put(dash, r, c1 + 1, c2, '=' + content, size=10)
            put(dash, r + 1, c1 + 1, c2, f'=IF({s}!C{src}="","","* "&SUBSTITUTE({s}!C{src},CHAR(10)," / "))', size=8.5)
        refs = '&'.join(f'IF({s}!D{j}="","","※ "&SUBSTITUTE({s}!D{j},CHAR(10)," ※ ")&"   ")' for j in range(2, 2 + REF_ROWS))
        over = f'IF(COUNTA({s}!A{cap + 2}:C200)>0,"▲ 칸이 모자라 {cap}개 뒤 일정은 빠졌습니다","")'
        put(dash, rr, c1, c2, f'={refs}&{over}', size=9)
        for c in range(c1, c2 + 1):
            cell = dash.cell(rr, c)
            cell.border = Border(top=DOT)
        outline(dash, c1, hr, c2, rr)

    dash.sheet_view.showGridLines = False
    dash.print_area = f'A1:L{LAST_ROW}'
    dash.page_setup.paperSize = dash.PAPERSIZE_A4
    dash.page_setup.orientation = 'portrait'
    dash.sheet_properties.pageSetUpPr.fitToPage = True
    dash.page_setup.fitToWidth = 1
    dash.page_setup.fitToHeight = 1
    dash.page_margins = PageMargins(left=0.39, right=0.39, top=0.39, bottom=0.39, header=0, footer=0)
    dash.print_options.horizontalCentered = True
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(OUT)
    print('saved', OUT)


if __name__ == '__main__':
    main()

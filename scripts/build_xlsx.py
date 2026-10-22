"""성경 애니메이션 제작표(엑셀) 생성 스크립트.

data/*.txt 의 분할 데이터를 읽어 제작표.xlsx 를 만든다.
실행: python3 scripts/build_xlsx.py
"""
import glob
import os
import re
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "성경애니메이션_제작표.xlsx")

CHAPTERS = {
    "창세기": 50, "출애굽기": 40, "레위기": 27, "민수기": 36, "신명기": 34,
    "여호수아": 24, "사사기": 21, "룻기": 4, "사무엘상": 31, "사무엘하": 24,
    "열왕기상": 22, "열왕기하": 25, "역대상": 29, "역대하": 36, "에스라": 10,
    "느헤미야": 13, "에스더": 10, "욥기": 42, "시편": 150, "잠언": 31,
    "전도서": 12, "아가": 8, "이사야": 66, "예레미야": 52, "예레미야애가": 5,
    "에스겔": 48, "다니엘": 12, "호세아": 14, "요엘": 3, "아모스": 9,
    "오바댜": 1, "요나": 4, "미가": 7, "나훔": 3, "하박국": 3, "스바냐": 3,
    "학개": 2, "스가랴": 14, "말라기": 4,
    "마태복음": 28, "마가복음": 16, "누가복음": 24, "요한복음": 21,
    "사도행전": 28, "로마서": 16, "고린도전서": 16, "고린도후서": 13,
    "갈라디아서": 6, "에베소서": 6, "빌립보서": 4, "골로새서": 4,
    "데살로니가전서": 5, "데살로니가후서": 3, "디모데전서": 6, "디모데후서": 4,
    "디도서": 3, "빌레몬서": 1, "히브리서": 13, "야고보서": 5, "베드로전서": 5,
    "베드로후서": 3, "요한일서": 5, "요한이서": 1, "요한삼서": 1, "유다서": 1,
    "요한계시록": 22,
}

# 전체를 연속으로 다루지 않고 핵심을 골라 다루는 책
SELECTIVE = {
    "시편": "150편 중 대표 시편 31편을 5권 구조 순서로 선별",
    "마가복음": "마태복음과 겹치는 사건은 제외, 마가 고유 장면 위주 선별",
    "누가복음": "탄생 이야기와 누가 고유 비유·장면 위주 선별",
    "요한복음": "21장 전체를 다루되 일부 강화(講話)는 대표 단락으로 압축",
    "역대상": "사무엘서와 겹치는 내용은 압축, 역대기 고유 관점 중심",
    "역대하": "열왕기와 겹치는 내용은 압축, 역대기 고유 관점 중심",
}

FONT = "Arial"
C_HEAD = "1F3864"
C_SUB = "D9E1F2"
C_INPUT = "FFF2CC"
C_OT = "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(bold=False, size=10, color="000000"):
    return Font(name=FONT, bold=bold, size=size, color=color)


WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ---------------------------------------------------------------- 데이터 읽기
def load():
    books, bs, rows = [], [], []
    book = None
    cur_b = None
    for path in sorted(glob.glob(os.path.join(ROOT, "data", "*.txt"))):
        for raw in open(path, encoding="utf-8"):
            line = raw.rstrip("\n")
            if not line.strip():
                continue
            if line.startswith("#"):
                parts = line[1:].strip().split("|")
                if len(parts) == 5 and ":" not in parts[0]:
                    name, abbr, testament, genre, default = [p.strip() for p in parts]
                    book = dict(name=name, abbr=abbr, testament=testament,
                                genre=genre, default=default)
                    books.append(book)
                continue
            if line.startswith("B "):
                p = line[2:].split("|")
                if len(p) not in (4, 5):
                    sys.exit(f"B 형식 오류: {line}")
                cur_b = dict(code=f"B{len(bs) + 1:03d}", title=p[0], intro=p[1],
                             christ=p[2], doctrine=p[3],
                             note=p[4] if len(p) == 5 else "", books=[])
                bs.append(cur_b)
                continue
            if line.startswith("A"):
                thematic = line.startswith("A* ")
                p = line[3:].split("|") if thematic else line[2:].split("|")
                if len(p) < 7:
                    sys.exit(f"A 형식 오류: {line}")
                p += [""] * (8 - len(p))
                rng, title, hook, kv, msg, people, place, caution = [x.strip() for x in p[:8]]
                if book["name"] not in cur_b["books"]:
                    cur_b["books"].append(book["name"])
                if book["name"] == "시편":
                    disp = f"시편 {rng}편"
                else:
                    disp = f"{book['abbr']} {rng}"
                rows.append(dict(
                    code=f"A{len(rows) + 1:04d}", b=cur_b, book=book, rng=rng,
                    disp=disp, kind="주제" if thematic else book["default"],
                    title=title, hook=hook, kv=f"{book['abbr']} {kv}", msg=msg,
                    people=people, place=place, caution=caution))
                continue
            sys.exit(f"알 수 없는 줄: {line}")
    return books, bs, rows


def chapters_of(rng):
    """'1:26-2:3' -> {1,2},  '5:1-32' -> {5},  '23' -> {23}"""
    if ":" not in rng:
        return {int(rng)}
    a, _, b = rng.partition("-")
    c1 = int(a.split(":")[0])
    c2 = int(b.split(":")[0]) if ":" in b else c1
    return set(range(c1, c2 + 1))


def span_text(first, lastr):
    """B타입 본문 범위: 첫 A의 시작 ~ 마지막 A의 끝 (예: 창 1:1~3:24)."""
    def ends(r):
        a, _, b = r["rng"].partition("-")
        if not b:
            return a, a
        if ":" in a and ":" not in b:
            b = a.split(":")[0] + ":" + b
        return a, b
    start, end = ends(first)[0], ends(lastr)[1]
    if first["book"]["name"] == "시편":
        start = f"시편 {start}"
        end = f"{end}편" if lastr["book"] is first["book"] else end
        if lastr["book"] is first["book"]:
            return f"{start}~{end}"
    else:
        start = f"{first['book']['abbr']} {start}"
    if lastr["book"] is not first["book"]:
        end = f"{lastr['book']['abbr']} {end}"
    return f"{start}~{end}"


def check(books, bs, rows):
    problems = []
    for bk in books:
        if bk["name"] not in CHAPTERS:
            problems.append(f"장 수 미등록: {bk['name']}")
            continue
        covered = set()
        for r in rows:
            if r["book"] is bk:
                covered |= chapters_of(r["rng"])
        missing = sorted(set(range(1, CHAPTERS[bk["name"]] + 1)) - covered)
        extra = sorted(c for c in covered if c > CHAPTERS[bk["name"]])
        if extra:
            problems.append(f"{bk['name']}: 존재하지 않는 장 {extra}")
        if missing and bk["name"] not in SELECTIVE:
            problems.append(f"{bk['name']}: 빠진 장 {missing}")
    if len(books) != 66:
        problems.append(f"책 수 {len(books)}권 (66권이어야 함)")
    for b in bs:
        n = sum(1 for r in rows if r["b"] is b)
        if n == 0:
            problems.append(f"{b['code']} 비어 있음")
    return problems


# ---------------------------------------------------------------- 공통 서식
def header(ws, row, titles, widths=None):
    for i, t in enumerate(titles, 1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = f(True, 10, "FFFFFF")
        c.fill = PatternFill("solid", fgColor=C_HEAD)
        c.alignment = CENTER
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30


def body(ws, r1, r2, c2, center_cols=()):
    for row in ws.iter_rows(min_row=r1, max_row=r2, max_col=c2):
        for c in row:
            c.font = f()
            c.border = BORDER
            c.alignment = CENTER if c.column in center_cols else WRAP


def title_block(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = f(True, 16, C_HEAD)
    ws["A2"] = sub
    ws["A2"].font = f(False, 10, "595959")


def section(ws, row, text, ncol=6):
    c = ws.cell(row=row, column=1, value=text)
    c.font = f(True, 12, "FFFFFF")
    for i in range(1, ncol + 1):
        ws.cell(row=row, column=i).fill = PatternFill("solid", fgColor=C_HEAD)
    return row + 1


def table(ws, row, titles, data, ncol=None, wrap_rows=True):
    ncol = ncol or len(titles)
    for i, t in enumerate(titles, 1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = f(True)
        c.fill = PatternFill("solid", fgColor=C_SUB)
        c.border = BORDER
        c.alignment = CENTER
    for d in data:
        row += 1
        for i, v in enumerate(d, 1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = f()
            c.border = BORDER
            c.alignment = WRAP
    return row + 2


# ---------------------------------------------------------------- 시트 작성
def build(books, bs, rows):
    wb = Workbook()
    n_a = len(rows)
    last = n_a + 1  # A타입 시트의 마지막 데이터 행

    # ===== 1. 안내 =====
    ws = wb.active
    ws.title = "안내"
    title_block(ws, "성경 애니메이션 시리즈 제작표 (대한예수교장로회 합동측 기준)",
                f"창세기~요한계시록 66권 장면·주제 분할 | 작성일 {date.today():%Y-%m-%d} | 개역개정 표기 기준")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 60

    r = section(ws, 4, "제작 변수 (노란 칸만 수정하세요)", 4)
    params = [
        ("시리즈 이름(가칭)", "1분 성경", "참고 영상 분석 후 확정. A타입 시트의 유튜브 제목안에 자동 반영"),
        ("숏폼 1편 길이(초)", 60, "A타입 업로드본 길이. 권장 45~60초"),
        ("롱폼에 넣는 A 1편 길이(초)", 52, "B타입으로 묶을 때 다음 편 예고(약 8초)를 뺀 길이"),
        ("B 인트로+아웃트로+연결(초)", 120, "진행 캐릭터 인트로 30초 + 중간 질문 10초 + 아웃트로 60초 + 연결 20초"),
        ("숏폼 주당 업로드 편수", 5, "업로드 예정일 계산에 사용 (월~금 기준)"),
        ("업로드 시작일", date(2027, 1, 4), "예시 날짜입니다. 실제 시작할 월요일로 바꾸세요"),
    ]
    prow = {}
    for label, val, note in params:
        ws.cell(row=r, column=1, value=label).font = f(True)
        c = ws.cell(row=r, column=3, value=val)
        c.fill = PatternFill("solid", fgColor=C_INPUT)
        c.font = f(False, 10, "0000FF")
        c.border = BORDER
        if isinstance(val, date):
            c.number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=4, value=note).font = f(False, 9, "595959")
        prow[label] = r
        r += 1
    P = {k: f"'안내'!$C${v}" for k, v in prow.items()}

    r = section(ws, r + 1, "전체 규모 (자동 계산)", 4)
    stats = [
        ("A타입(숏폼) 총 편수", f"=COUNTA('A타입_제작표'!$B$2:$B${last})", "편"),
        ("B타입(롱폼) 총 편수", f"=COUNTA('B타입_롱폼'!$B$2:$B${len(bs) + 1})", "편"),
        ("구약 A타입", f"=COUNTIF('A타입_제작표'!$F$2:$F${last},\"구약\")", "편"),
        ("신약 A타입", f"=COUNTIF('A타입_제작표'!$F$2:$F${last},\"신약\")", "편"),
        ("숏폼 총 분량", f"=ROUND(C{r}*{P['숏폼 1편 길이(초)']}/3600,1)", "시간"),
        ("롱폼 총 분량", f"=ROUND(SUM('B타입_롱폼'!$G$2:$G${len(bs) + 1})/3600,1)", "시간"),
        ("숏폼 완주 기간", f"=ROUNDUP(C{r}/{P['숏폼 주당 업로드 편수']},0)", "주"),
        ("완주 예상일(마지막 숏폼)", f"=MAX('A타입_제작표'!$S$2:$S${last})", ""),
    ]
    stat_start = r
    for i, (label, formula, unit) in enumerate(stats):
        ws.cell(row=r, column=1, value=label).font = f(True)
        c = ws.cell(row=r, column=3, value=formula)
        c.font = f()
        c.border = BORDER
        if "예상일" in label:
            c.number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=4, value=unit).font = f(False, 9, "595959")
        r += 1
    # 숏폼 총 분량·완주 기간은 첫 줄(총 편수)을 참조
    ws.cell(row=stat_start + 4, column=3).value = f"=ROUND(C{stat_start}*{P['숏폼 1편 길이(초)']}/3600,1)"
    ws.cell(row=stat_start + 6, column=3).value = f"=ROUNDUP(C{stat_start}/{P['숏폼 주당 업로드 편수']},0)"

    r = section(ws, r + 1, "분할 기준", 4)
    rules = [
        ("A타입 (숏폼·릴스)", "45~60초, 장면 1개 · 메시지 1개. 이야기 책은 '장면형', 시가서·서신서·예언 등은 '주제형'으로 분할"),
        ("B타입 (유튜브 롱폼)", "A타입 7~14편을 하나의 이야기 단위로 묶어 약 8~14분. 여러 권을 묶은 B도 있음(예: 룻기+사무엘상 1~3장, 소선지서)"),
        ("진행 캐릭터 (절충안)", "A타입: 마지막 2~3초 등장(다음 편 안내) / B타입: 인트로·중간 질문·아웃트로 진행"),
        ("스토리형 원칙", "쉬운 말 내레이션 + 핵심 구절 1개만 개역개정 인용. 성경에 없는 대사·장면은 최소화하고 '상상 장면'으로 감수"),
        ("구속사 관점", "모든 B타입에 '그리스도 연결점'을 두어 도덕 교훈에 그치지 않고 그리스도를 가리키도록 구성"),
        ("선별 분할한 책", "; ".join(f"{k}: {v}" for k, v in SELECTIVE.items())),
    ]
    for label, text in rules:
        ws.cell(row=r, column=1, value=label).font = f(True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=2, value=text)
        c.font = f()
        c.alignment = WRAP
        ws.row_dimensions[r].height = 30 if len(text) < 90 else 60
        r += 1

    r = section(ws, r + 1, "시트 구성", 4)
    sheets = [
        ("권별요약", "66권별 A·B 편수, 분량, 회차 범위"),
        ("B타입_롱폼", "롱폼 목록: 진행자 인트로 질문, 그리스도 연결점, 핵심 교리, 공개 예정일, 제작 상태"),
        ("A타입_제작표", "숏폼 전체 목록: 본문, 제목, 첫 3초 훅, 핵심 구절, 메시지, 인물, 배경, 교리·연출 주의, 일정, 제작 상태"),
        ("인물_캐릭터시트", "주요 인물 외형·의상 가안과 삼위 하나님·천사·사탄 표현 규칙, 진행 캐릭터(가안)"),
        ("교리감수_기준", "합동측 신앙 표준, 핵심 교리 체크, 이단 해석 경계 목록, 감수 절차"),
        ("제작템플릿", "A·B 구성 공식, 프롬프트 5블록, 기술 규격, 메타데이터, 저작권·업로드 체크리스트"),
        ("절기특집", "교회 절기별 추천 편성"),
        ("표기통일", "개역개정 기준 용어·인명 표기"),
    ]
    for name, desc in sheets:
        ws.cell(row=r, column=1, value=name).font = f(True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=2, value=desc).font = f()
        r += 1
    ws.cell(row=r + 1, column=1, value="※ 노란 칸 = 수정 가능한 입력칸 / 상태 열은 목록에서 선택 / 그 외 숫자는 수식으로 자동 계산").font = f(False, 9, "C00000")

    # ===== 2. A타입_제작표 =====
    wa = wb.create_sheet("A타입_제작표")
    a_cols = ["순번", "A코드", "회차", "B코드", "B타입 제목", "구약/신약", "책", "본문", "유형",
              "제목", "첫 3초 훅", "핵심 구절", "한 줄 메시지", "등장인물", "장소·배경",
              "교리·연출 주의", "유튜브 제목안", "업로드 주차", "업로드 예정일",
              "대본", "교리감수", "영상", "업로드", "비고", "다음 편 예고"]
    widths = [6, 8, 8, 7, 22, 7, 10, 16, 6, 20, 30, 12, 32, 22, 22, 34, 40, 7, 11, 8, 8, 8, 8, 16, 22]
    header(wa, 1, a_cols, widths)
    for i, rw in enumerate(rows, 2):
        vals = [i - 1, rw["code"], f"#{i - 1:04d}", rw["b"]["code"], rw["b"]["title"],
                rw["book"]["testament"], rw["book"]["name"], rw["disp"], rw["kind"],
                rw["title"], rw["hook"], rw["kv"], rw["msg"], rw["people"], rw["place"],
                rw["caution"]]
        for j, v in enumerate(vals, 1):
            wa.cell(row=i, column=j, value=v)
        wa.cell(row=i, column=17, value=f"=\"[\"&{P['시리즈 이름(가칭)']}&\" \"&C{i}&\"] \"&J{i}&\" | \"&H{i}")
        wa.cell(row=i, column=18, value=f"=INT((A{i}-1)/{P['숏폼 주당 업로드 편수']})+1")
        wa.cell(row=i, column=19, value=f"={P['업로드 시작일']}+(R{i}-1)*7+MOD(A{i}-1,{P['숏폼 주당 업로드 편수']})")
        wa.cell(row=i, column=19).number_format = "yyyy-mm-dd"
        for j in range(20, 24):
            wa.cell(row=i, column=j, value="미착수")
        nxt = f"=\"다음 편: \"&J{i + 1}" if i < last else "시리즈 완결"
        wa.cell(row=i, column=25, value=nxt)
    body(wa, 2, last, len(a_cols), center_cols=(1, 2, 3, 4, 6, 7, 9, 18, 19, 20, 21, 22, 23))
    for i, rw in enumerate(rows, 2):
        if rw["book"]["testament"] == "구약":
            for j in (1, 2, 3, 4, 5, 6, 7, 8):
                wa.cell(row=i, column=j).fill = PatternFill("solid", fgColor=C_OT)
        if rw["caution"]:
            wa.cell(row=i, column=16).font = f(False, 10, "C00000")
    wa.freeze_panes = "K2"
    wa.auto_filter.ref = f"A1:{get_column_letter(len(a_cols))}{last}"
    dv = DataValidation(type="list", formula1='"미착수,진행중,완료,보류"', allow_blank=True)
    dv.add(f"T2:W{last}")
    wa.add_data_validation(dv)

    # ===== 3. B타입_롱폼 =====
    wbs = wb.create_sheet("B타입_롱폼")
    nb_last = len(bs) + 1
    b_cols = ["순번", "B코드", "B타입 제목", "포함 책", "본문 범위", "A 편수", "예상 길이(초)",
              "예상 길이(분)", "첫 A코드", "마지막 A코드", "진행자 인트로 질문(삶의 고민)",
              "그리스도 연결점", "핵심 교리", "B 전체 주의", "공개 예정일",
              "대본", "교리감수", "영상", "업로드", "비고"]
    header(wbs, 1, b_cols, [6, 7, 26, 18, 26, 7, 9, 9, 9, 9, 32, 40, 32, 30, 11, 8, 8, 8, 8, 16])
    for i, b in enumerate(bs, 2):
        mine = [r for r in rows if r["b"] is b]
        first, lastr = mine[0], mine[-1]
        span = span_text(first, lastr)
        vals = [i - 1, b["code"], b["title"], " · ".join(b["books"]), span]
        for j, v in enumerate(vals, 1):
            wbs.cell(row=i, column=j, value=v)
        wbs.cell(row=i, column=6, value=f"=COUNTIF('A타입_제작표'!$D$2:$D${last},B{i})")
        wbs.cell(row=i, column=7, value=f"=F{i}*{P['롱폼에 넣는 A 1편 길이(초)']}+{P['B 인트로+아웃트로+연결(초)']}")
        wbs.cell(row=i, column=8, value=f"=ROUND(G{i}/60,1)")
        wbs.cell(row=i, column=9, value=f"=INDEX('A타입_제작표'!$B$2:$B${last},MATCH(B{i},'A타입_제작표'!$D$2:$D${last},0))")
        wbs.cell(row=i, column=10, value=f"=INDEX('A타입_제작표'!$B$2:$B${last},MATCH(B{i},'A타입_제작표'!$D$2:$D${last},0)+F{i}-1)")
        for j, v in zip((11, 12, 13, 14), (b["intro"], b["christ"], b["doctrine"], b["note"])):
            wbs.cell(row=i, column=j, value=v)
        wbs.cell(row=i, column=15, value=f"=_xlfn.MAXIFS('A타입_제작표'!$S$2:$S${last},'A타입_제작표'!$D$2:$D${last},B{i})+1")
        wbs.cell(row=i, column=15).number_format = "yyyy-mm-dd"
        for j in range(16, 20):
            wbs.cell(row=i, column=j, value="미착수")
    body(wbs, 2, nb_last, len(b_cols), center_cols=(1, 2, 6, 7, 8, 9, 10, 15, 16, 17, 18, 19))
    wbs.freeze_panes = "D2"
    wbs.auto_filter.ref = f"A1:{get_column_letter(len(b_cols))}{nb_last}"
    dvb = DataValidation(type="list", formula1='"미착수,진행중,완료,보류"', allow_blank=True)
    dvb.add(f"P2:S{nb_last}")
    wbs.add_data_validation(dvb)
    wbs.cell(row=1, column=15).comment = Comment("해당 B에 포함된 마지막 A타입 공개 다음 날", "제작표")
    wbs.cell(row=1, column=7).comment = Comment("A 편수 × 롱폼용 A 길이 + 인트로·아웃트로·연결 (안내 시트 변수)", "제작표")

    # ===== 4. 권별요약 =====
    wk = wb.create_sheet("권별요약")
    k_cols = ["순서", "책", "약어", "구약/신약", "장르", "장 수", "A 편수", "B 편수(포함)",
              "숏폼 분량(분)", "시작 회차", "끝 회차", "분할 방식"]
    header(wk, 1, k_cols, [6, 14, 6, 9, 10, 7, 8, 10, 10, 9, 9, 50])
    rngA = f"'A타입_제작표'!$G$2:$G${last}"
    rngB = f"'A타입_제작표'!$D$2:$D${last}"
    rngN = f"'A타입_제작표'!$A$2:$A${last}"
    for i, bk in enumerate(books, 2):
        vals = [i - 1, bk["name"], bk["abbr"], bk["testament"], bk["genre"], CHAPTERS[bk["name"]]]
        for j, v in enumerate(vals, 1):
            wk.cell(row=i, column=j, value=v)
        wk.cell(row=i, column=7, value=f"=COUNTIF({rngA},B{i})")
        wk.cell(row=i, column=8, value=f"=SUMPRODUCT(({rngA}=B{i})/COUNTIFS({rngA},{rngA},{rngB},{rngB}))")
        wk.cell(row=i, column=9, value=f"=ROUND(G{i}*{P['숏폼 1편 길이(초)']}/60,1)")
        wk.cell(row=i, column=10, value=f"=_xlfn.MINIFS({rngN},{rngA},B{i})")
        wk.cell(row=i, column=11, value=f"=_xlfn.MAXIFS({rngN},{rngA},B{i})")
        mode = SELECTIVE.get(bk["name"], "전체 장을 순서대로 빠짐없이 분할")
        wk.cell(row=i, column=12, value=mode)
    tot = len(books) + 2
    wk.cell(row=tot, column=2, value="합계")
    wk.cell(row=tot, column=6, value=f"=SUM(F2:F{tot - 1})")
    wk.cell(row=tot, column=7, value=f"=SUM(G2:G{tot - 1})")
    wk.cell(row=tot, column=9, value=f"=SUM(I2:I{tot - 1})")
    body(wk, 2, tot, len(k_cols), center_cols=(1, 3, 4, 5, 6, 7, 8, 9, 10, 11))
    for j in range(1, len(k_cols) + 1):
        wk.cell(row=tot, column=j).font = f(True)
        wk.cell(row=tot, column=j).fill = PatternFill("solid", fgColor=C_SUB)
    wk.freeze_panes = "C2"
    wk.cell(row=1, column=8).comment = Comment("그 책의 장면이 들어간 B타입 수. 여러 권을 묶은 B는 각 책에 모두 집계됨", "제작표")

    # ===== 5~9 참고 시트 =====
    from sheets_ref import characters, doctrine, templates, seasons, terms
    characters(wb, rows, last, f, header, body, title_block, section, table, C_INPUT, BORDER, WRAP)
    doctrine(wb, f, title_block, section, table)
    templates(wb, f, title_block, section, table)
    seasons(wb, rows, f, header, body, title_block)
    terms(wb, f, title_block, section, table)

    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    books, bs, rows = load()
    problems = check(books, bs, rows)
    print(f"책 {len(books)}권 / B {len(bs)}편 / A {len(rows)}편")
    for p in problems:
        print("확인 필요:", p)
    if any("빠진 장" in p or "존재하지" in p or "66권" in p for p in problems):
        sys.exit(1)
    print("저장:", build(books, bs, rows))

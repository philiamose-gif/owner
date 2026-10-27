"""프로젝트2 말씀 숏폼 제작표 생성.

실행: python3 project2/build.py
"""
import os
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))
import build_xlsx as bx  # noqa: E402  (서식 도구, 66권 정보 재사용)
import verify_bible as vb  # noqa: E402  (역본 본문)

OUT = os.path.join(HERE, "말씀숏폼_제작표.xlsx")
PER_DAY = 2  # 트랙별 하루 편수
TARGET_DAYS = 182  # 6개월

# 주제 → 화면 배경 분류 (사람 얼굴·하나님 형상 없이 풍경과 빛으로만)
BACKGROUNDS = {
    "새벽빛": ("해 뜨는 지평선, 구름 사이로 퍼지는 새벽빛",
              ["빛", "소망", "부활", "새 마음", "생명", "회복", "하나님 나라", "기쁨"]),
    "잔잔한 물": ("잔잔한 호수, 맑은 시냇물, 물 위의 햇빛",
               ["평안", "쉼", "위로", "치료", "은혜", "용서"]),
    "바위산·광야": ("바위산, 광야 길, 견고한 바위와 넓은 하늘",
                ["믿음", "신뢰", "담대함", "능력", "승리", "보호", "피난처", "연단", "신실하심"]),
    "들판·나무": ("밀밭, 올리브 나무, 열매 맺은 가지, 씨앗과 새싹",
               ["말씀", "지혜", "성장", "나눔", "섬김", "공급하심", "복", "축복", "감사", "정의", "경건"]),
    "하늘·별": ("별이 가득한 밤하늘, 구름 위 빛줄기, 은하수",
              ["창조", "하나님", "섭리", "거룩", "찬양", "예배", "경외", "임재", "성령"]),
    "길·등불": ("굽이진 오솔길, 밤길을 밝히는 등불, 먼 곳으로 이어진 길",
              ["인도", "동행", "순종", "부르심", "헌신", "결단"]),
    "창가·식탁": ("따뜻한 햇살이 드는 창가, 빵과 잔이 놓인 나무 식탁 (사람 없이)",
               ["가정", "사랑", "연합", "교회", "마음", "말", "겸손"]),
    "제단·십자가": ("돌 제단, 언덕 위 십자가 실루엣과 빛, 열린 두루마리",
                ["기도", "회개", "언약", "속죄", "구원", "복음", "자유", "계명"]),
}
THEME_BG = {t: k for k, (_, ts) in BACKGROUNDS.items() for t in ts}
f, header, body, section, table = bx.f, bx.header, bx.body, bx.section, bx.table
C_INPUT, C_SUB, BORDER = bx.C_INPUT, bx.C_SUB, bx.BORDER
STATUS = '"미착수,진행중,완료,보류"'


def load():
    books = {b["name"]: b for b in bx.load()[0]}
    verses = []
    for track, fname in (("구약", "구약.txt"), ("신약", "신약.txt")):
        n = 0
        for raw in open(os.path.join(HERE, "data", fname), encoding="utf-8"):
            line = raw.rstrip("\n")
            if not line.strip() or line.startswith("#"):
                continue
            p = line.split("|")
            if len(p) != 4:
                sys.exit(f"형식 오류: {line}")
            name, ref, theme, note = p
            if name not in books or books[name]["testament"] != track:
                sys.exit(f"책 이름 오류({track}): {line}")
            ch = int(ref.split(":")[0])
            if not 1 <= ch <= bx.CHAPTERS[name]:
                sys.exit(f"없는 장: {line}")
            if theme not in THEME_BG:
                sys.exit(f"배경이 정해지지 않은 주제: {theme}")
            n += 1
            verses.append(dict(track=track, no=n, book=books[name], ref=ref, theme=theme, note=note,
                               disp=f"{books[name]['abbr']} {ref}"))
    return list(books.values()), verses


def check(books, verses):
    problems = []
    for t in ("구약", "신약"):
        tv = [v for v in verses if v["track"] == t]
        if len(tv) != PER_DAY * TARGET_DAYS:
            problems.append(f"{t} {len(tv)}편 (목표 {PER_DAY * TARGET_DAYS}편)")
        order = [list(books).index(v["book"]) for v in tv]
        if order != sorted(order):
            problems.append(f"{t} 성경 순서가 뒤섞임")
    for b in books:
        if not any(v["book"] is b for v in verses):
            problems.append(f"{b['name']} 구절 없음")
    keys = [v["disp"] for v in verses]
    dup = sorted({k for k in keys if keys.count(k) > 1})
    if dup:
        problems.append(f"중복 구절: {dup}")
    return problems


def build(books, verses):
    wb = Workbook()
    n = len(verses)
    last = n + 1
    V = lambda c: f"'구절목록'!${c}$2:${c}${last}"

    # ===== 안내 =====
    ws = wb.active
    ws.title = "안내"
    bx.title_block(ws, "프로젝트2 · 말씀 숏폼 제작표",
                   f"성경 66권 핵심 구절 {n}편을 6개월에 · 구약/신약 트랙 · 성경 본문: 개역한글판 | 작성일 {date.today():%Y-%m-%d}")
    for c, w in zip("ABCD", (30, 18, 16, 64)):
        ws.column_dimensions[c].width = w
    r = section(ws, 4, "제작 변수 (노란 칸만 수정하세요)", 4)
    params = [
        ("말씀 시리즈 이름(가칭)", "오늘의 말씀", "같은 채널 안의 재생목록 이름. 제목·해시태그에 반영"),
        ("시작일", date(2027, 1, 1), "예시 날짜입니다. 프로젝트1과 상관없이 정하면 됩니다"),
        ("트랙별 하루 편수", PER_DAY, "구약·신약 각각. 2편이면 하루 총 4편"),
        ("구약 1번째 업로드 시간", "06:30", "출근·등교 전 묵상 시간"),
        ("신약 1번째 업로드 시간", "12:30", "점심 시간"),
        ("구약 2번째 업로드 시간", "18:30", "퇴근 시간"),
        ("신약 2번째 업로드 시간", "21:30", "잠들기 전. 프로젝트1(애니메이션)은 17:00 전후로 두어 겹치지 않게"),
        ("낭독 속도 (글자/초)", 5, "차분한 낭독 기준. 권장 영상 길이 계산에 사용"),
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
    for label, formula, unit, fmt in [
        ("전체 편수", f"=COUNTA({V('B')})", "편", None),
        ("구약 / 신약", f"=COUNTIF({V('C')},\"구약\")&\" / \"&COUNTIF({V('C')},\"신약\")", "편", None),
        ("포함된 책", "=COUNTIF('권별점검'!$C$2:$C$67,\">0\")", "권 / 66권", None),
        ("마지막 업로드일", f"=MAX({V('R')})", "", "date"),
        ("전체 기간", f"=MAX({V('R')})-{P['시작일']}+1", "일 (목표 182일 이내)", None),
        ("본문 자동 입력", f"=COUNTIF({V('L')},\"자동 입력*\")+COUNTIF({V('L')},\"대조 완료\")", "편 (개역한글 원문에서)", None),
        ("띄어쓰기 대조가 필요한 구절", f"=COUNTIF({V('L')},\"*대조 필요\")", "편 (여러 줄에 걸친 구절)", None),
    ]:
        ws.cell(row=r, column=1, value=label).font = f(True)
        c = ws.cell(row=r, column=3, value=formula)
        c.border = BORDER
        if fmt == "date":
            c.number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=4, value=unit).font = f(False, 9, "595959")
        r += 1

    r = section(ws, r + 1, "운영 원칙", 4)
    for label, text in [
        ("말씀에만 집중", "영상에는 해설·진행자·내레이션 없이 구절 낭독과 자막만. 설명이 꼭 필요한 구절은 '설명란 문구'를 유튜브 설명란·인스타 게시글에만 적음"),
        ("트랙", "구약 트랙과 신약 트랙이 각각 성경 순서대로 동시에 진행 (창세기→말라기, 마태복음→요한계시록)"),
        ("채널", "프로젝트1과 같은 채널. '말씀' 재생목록으로 구분하고 업로드 시간을 겹치지 않게"),
        ("성경 본문", "개역한글판만 사용. 본문은 reference/krv.tsv(개역한글 원문 변환본)에서 자동으로 넣음. '대조 필요' 표시 구절은 띄어쓰기를 원본 PDF와 한 번 더 확인"),
        ("선정 기준", "66권 모두 1구절 이상. 책의 분량과 많이 암송·인용되는 구절을 함께 고려. 한 절만 떼어 쓰면 뜻이 달라지는 구절은 피하거나 설명란 문구를 붙임"),
    ]:
        ws.cell(row=r, column=1, value=label).font = f(True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=2, value=text)
        c.font = f()
        c.alignment = bx.WRAP
        ws.row_dimensions[r].height = 30
        r += 1

    # ===== 구절목록 =====
    wv = wb.create_sheet("구절목록")
    cols = ["순번", "코드", "트랙", "트랙 회차", "회차 표기", "책", "장절", "자막 출처 표기", "주제", "화면 배경",
            "본문 (개역한글)", "본문 상태", "글자 수", "권장 길이(초)", "제목안", "해시태그", "설명란 문구",
            "업로드일", "업로드 시간", "요일", "영상", "업로드", "조회수", "시청 vs 넘김(%)", "비고",
            "확인 표시", "WEB (영어, 확장용)"]
    header(wv, 1, cols, [6, 8, 6, 7, 13, 11, 9, 16, 10, 11, 50, 14, 7, 9, 44, 40, 50, 11, 9, 5, 8, 8, 9, 9, 16, 24, 50])
    for i, v in enumerate(verses, 2):
        krv, flags = vb.text(v["book"]["abbr"], v["ref"])  # 시편 표제는 말씀 카드에서 생략
        status = "자동 입력 · 대조 필요" if flags & vb.REVIEW else "자동 입력"
        web = vb.text(v["book"]["abbr"], v["ref"], vb.WEB)[0] if vb.WEB else None
        vals = [i - 1, f"W{i - 1:04d}", v["track"], v["no"], None, v["book"]["name"], v["ref"],
                f"({v['disp']}, 개역한글)", v["theme"], THEME_BG[v["theme"]], krv, status if krv else "미입력"]
        wv.cell(row=i, column=26, value=", ".join(sorted(flags)))
        wv.cell(row=i, column=27, value=web or "")
        for j, val in enumerate(vals, 1):
            wv.cell(row=i, column=j, value=val)
        wv.cell(row=i, column=5, value=f"=C{i}&\"말씀 #\"&TEXT(D{i},\"000\")")
        wv.cell(row=i, column=13, value=f"=LEN(K{i})")
        wv.cell(row=i, column=14, value=f"=IF(M{i}=0,\"-\",ROUND(3+M{i}/{P['낭독 속도 (글자/초)']},0))")
        wv.cell(row=i, column=15, value=(
            f"=IF(K{i}=\"\",\"(본문 입력 후 자동) | {v['disp']}\",IF(LEN(K{i})>28,LEFT(K{i},28)&\"…\",K{i})&\" | {v['disp']}\")"))
        tag_theme = v["theme"].replace(" ", "")
        wv.cell(row=i, column=16, value=(
            f"=\"#\"&SUBSTITUTE({P['말씀 시리즈 이름(가칭)']},\" \",\"\")&\" #성경구절 #말씀 #{v['book']['name']} #{tag_theme}\""))
        footer = "성경 본문: 개역한글판"
        wv.cell(row=i, column=17, value=f"{v['note']}\n{footer}" if v["note"] else footer)
        wv.cell(row=i, column=18, value=f"={P['시작일']}+INT((D{i}-1)/{P['트랙별 하루 편수']})")
        wv.cell(row=i, column=18).number_format = "yyyy-mm-dd"
        first, second = (("구약 1번째 업로드 시간", "구약 2번째 업로드 시간") if v["track"] == "구약"
                         else ("신약 1번째 업로드 시간", "신약 2번째 업로드 시간"))
        wv.cell(row=i, column=19, value=f"=IF(MOD(D{i}-1,{P['트랙별 하루 편수']})=0,{P[first]},{P[second]})")
        wv.cell(row=i, column=20, value=f"=CHOOSE(WEEKDAY(R{i},2),\"월\",\"화\",\"수\",\"목\",\"금\",\"토\",\"일\")")
        wv.cell(row=i, column=21, value="미착수")
        wv.cell(row=i, column=22, value="미착수")
    body(wv, 2, last, len(cols), center_cols=(1, 2, 3, 4, 5, 6, 7, 9, 10, 12, 13, 14, 18, 19, 20, 21, 22, 23, 24))
    for i, v in enumerate(verses, 2):
        for j in (11, 23, 24):
            wv.cell(row=i, column=j).fill = PatternFill("solid", fgColor=C_INPUT)
        if v["track"] == "구약":
            for j in (1, 2, 3, 4, 5):
                wv.cell(row=i, column=j).fill = PatternFill("solid", fgColor=bx.C_OT)
        if v["note"]:
            wv.cell(row=i, column=17).font = f(False, 10, "C00000")
    wv.freeze_panes = "I2"
    wv.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{last}"
    dv1 = DataValidation(type="list", formula1='"미입력,자동 입력,자동 입력 · 대조 필요,대조 완료"', allow_blank=True)
    dv1.add(f"L2:L{last}")
    dv2 = DataValidation(type="list", formula1=STATUS, allow_blank=True)
    dv2.add(f"U2:V{last}")
    wv.add_data_validation(dv1)
    wv.add_data_validation(dv2)

    # ===== 권별점검 =====
    wk = wb.create_sheet("권별점검")
    header(wk, 1, ["책", "구약/신약", "구절 수", "장 수", "첫 업로드일", "마지막 업로드일", "결과"],
           [14, 9, 8, 7, 12, 14, 10])
    for i, b in enumerate(books, 2):
        wk.cell(row=i, column=1, value=b["name"])
        wk.cell(row=i, column=2, value=b["testament"])
        wk.cell(row=i, column=3, value=f"=COUNTIF({V('F')},A{i})")
        wk.cell(row=i, column=4, value=bx.CHAPTERS[b["name"]])
        wk.cell(row=i, column=5, value=f"=_xlfn.MINIFS({V('R')},{V('F')},A{i})")
        wk.cell(row=i, column=6, value=f"=_xlfn.MAXIFS({V('R')},{V('F')},A{i})")
        for j in (5, 6):
            wk.cell(row=i, column=j).number_format = "yyyy-mm-dd"
        wk.cell(row=i, column=7, value=f"=IF(C{i}>0,\"정상\",\"구절 없음\")")
    body(wk, 2, len(books) + 1, 7, center_cols=(2, 3, 4, 5, 6, 7))
    wk.freeze_panes = "B2"

    # ===== 누락점검 =====
    wc = wb.create_sheet("누락점검")
    bx.title_block(wc, "누락 점검", "결과가 모두 '정상'이어야 합니다.")
    for c, w in zip("ABCD", (44, 14, 14, 12)):
        wc.column_dimensions[c].width = w
    items = [
        ("전체 편수", f"=COUNTA({V('B')})", n),
        ("구약 편수", f"=COUNTIF({V('C')},\"구약\")", PER_DAY * TARGET_DAYS),
        ("신약 편수", f"=COUNTIF({V('C')},\"신약\")", PER_DAY * TARGET_DAYS),
        ("구절이 없는 책", "=COUNTIF('권별점검'!$C$2:$C$67,0)", 0),
        ("포함된 책 수", "=COUNTIF('권별점검'!$C$2:$C$67,\">0\")", 66),
        ("중복 구절", f"=SUMPRODUCT((COUNTIF({V('H')},{V('H')})>1)*1)", 0),
        ("같은 날·같은 시간 겹친 업로드", f"=SUMPRODUCT((COUNTIFS({V('R')},{V('R')},{V('S')},{V('S')})>1)*1)", 0),
        ("6개월(182일) 안에 끝나는가 (1 = 예)", f"=IF(MAX({V('R')})-{P['시작일']}+1<=182,1,0)", 1),
        ("트랙 회차 중복", f"=SUMPRODUCT((COUNTIFS({V('C')},{V('C')},{V('D')},{V('D')})>1)*1)", 0),
        ("화면 배경이 정해지지 않은 구절", f"=COUNTBLANK({V('J')})", 0),
    ]
    r = section(wc, 4, "점검 항목", 4)
    for j, t in enumerate(["항목", "현재 값", "기대 값", "결과"], 1):
        c = wc.cell(row=r, column=j, value=t)
        c.font = f(True)
        c.fill = PatternFill("solid", fgColor=C_SUB)
    first = r + 1
    for k, (label, cur, exp) in enumerate(items):
        rr = first + k
        wc.cell(row=rr, column=1, value=label)
        wc.cell(row=rr, column=2, value=cur)
        wc.cell(row=rr, column=3, value=exp)
        wc.cell(row=rr, column=4, value=f"=IF(B{rr}=C{rr},\"정상\",\"확인 필요\")")
        for j in range(1, 5):
            wc.cell(row=rr, column=j).border = BORDER
            wc.cell(row=rr, column=j).font = f(j == 4)
    rr = first + len(items)
    wc.cell(row=rr, column=1, value="전체 결과").font = f(True)
    wc.cell(row=rr, column=4, value=f"=IF(COUNTIF(D{first}:D{rr - 1},\"정상\")={len(items)},\"모두 정상\",\"확인 필요 있음\")").font = f(True, 11, "C00000")
    wc.cell(row=rr + 2, column=1, value="참고: 본문 입력 진행").font = f(True)
    wc.cell(row=rr + 3, column=1, value="본문이 아직 입력되지 않은 구절")
    wc.cell(row=rr + 3, column=2, value=f"=COUNTBLANK({V('K')})")
    wc.cell(row=rr + 4, column=1, value="본문 대조 완료")
    wc.cell(row=rr + 4, column=2, value=f"=COUNTIF({V('L')},\"대조 완료\")")

    # ===== 제작템플릿 =====
    wt = wb.create_sheet("제작템플릿")
    bx.title_block(wt, "말씀 숏폼 제작 템플릿", "말씀에만 집중: 해설·진행자 없이 구절과 낭독만")
    for c, w in zip("ABCD", (24, 60, 40, 30)):
        wt.column_dimensions[c].width = w
    r = section(wt, 4, "1. 영상 구성 (15~25초)", 4)
    r = table(wt, r, ["구간", "내용", "연결 칸", "비고"], [
        ("0~1초", "배경 화면이 열리며 시그니처 사운드(프로젝트1과 다른 짧은 소리)", "화면 배경", "첫 프레임부터 구절 첫 줄이 보이게 해 넘김을 줄임"),
        ("1초~끝 3초 전", "구절을 한 줄씩 자막으로 띄우며 차분히 낭독. 해설·배경 설명 없음", "본문 (개역한글)", "권장 길이(초) 칸 참고"),
        ("마지막 3초", "구절 전체 자막 + 출처 '(창 1:1, 개역한글)'", "자막 출처 표기", "멈춰서 다시 읽을 시간"),
        ("반복 재생", "마지막 화면에서 첫 화면으로 자연스럽게 이어지게", "-", "반복 시청 유도"),
    ])
    r = section(wt, r, "2. 화면 배경 기준 (사람 얼굴·하나님 형상 없이 풍경과 빛)", 4)
    r = table(wt, r, ["배경 분류", "장면", "해당 주제", "구절 수"],
              [(k, d, ", ".join(ts), f"=COUNTIF({V('J')},\"{k}\")") for k, (d, ts) in BACKGROUNDS.items()])
    r = section(wt, r, "3. 자막·음성", 4)
    r = table(wt, r, ["항목", "기준", "비고"], [
        ("자막", "굵은 명조 또는 고딕, 한 줄 12~14자, 화면 가운데", "개역한글 원문 그대로. 띄어쓰기·맞춤법을 현대식으로 고치지 않음"),
        ("낭독", "차분하고 또렷한 목소리 1명으로 고정", "AI 음성은 상업 이용 허용 여부 확인"),
        ("배경음악", "잔잔한 저작권 무료 음원. 낭독보다 충분히 작게", "찬송가·CCM 사용 금지"),
        ("화면 비율", "1080×1920 (9:16)", "상단·하단 플랫폼 UI 영역에는 자막을 두지 않음"),
    ])
    r = section(wt, r, "4. 설명란·게시글", 4)
    r = table(wt, r, ["항목", "형식", "비고"], [
        ("유튜브 제목", "구절 앞부분… | 장절", "'제목안' 칸 자동 생성"),
        ("설명란 / 인스타 게시글", "구절 전문 + 장절 + (필요한 구절만) 설명란 문구 + '성경 본문: 개역한글판'", "'설명란 문구' 칸이 빨간 글씨인 구절은 문맥 설명을 꼭 넣기"),
        ("해시태그", "#시리즈명 #성경구절 #말씀 #책이름 #주제", "'해시태그' 칸 자동 생성"),
    ])
    r = section(wt, r, "5. 저작권·교리 원칙", 4)
    table(wt, r, ["원칙", "내용"], [
        ("본문", "개역한글판만 사용(저작재산권 보호기간 만료). 개역개정·새번역 등 다른 번역 문장을 섞지 않음"),
        ("원문 확보", "공개된 개역한글 원문에서 장절별로 붙여 넣고 '본문 상태'를 '대조 완료'로 표시"),
        ("문맥", "한 절만 떼어 쓰면 뜻이 달라질 수 있는 구절은 설명란 문구를 붙임 (예: 빌 4:13, 렘 29:11)"),
        ("감수", "프로젝트1의 교리감수 기준(합동측)을 그대로 적용"),
    ])

    # ===== 성과기록 =====
    wp = wb.create_sheet("성과기록")
    bx.title_block(wp, "성과 기록", "구절목록의 노란 칸(조회수·시청 vs 넘김)에 입력하면 자동 집계됩니다.")
    header_row = 4
    for j, t in enumerate(["주제", "구절 수", "입력된 편", "평균 조회수", "평균 시청 vs 넘김(%)"], 1):
        c = wp.cell(row=header_row, column=j, value=t)
        c.font = f(True)
        c.fill = PatternFill("solid", fgColor=C_SUB)
    themes = sorted(THEME_BG)
    for k, t in enumerate(themes, header_row + 1):
        wp.cell(row=k, column=1, value=t)
        wp.cell(row=k, column=2, value=f"=COUNTIF({V('I')},A{k})")
        wp.cell(row=k, column=3, value=f"=COUNTIFS({V('I')},A{k},{V('W')},\"<>\")")
        wp.cell(row=k, column=4, value=f"=IFERROR(AVERAGEIFS({V('W')},{V('I')},A{k}),\"-\")")
        wp.cell(row=k, column=5, value=f"=IFERROR(AVERAGEIFS({V('X')},{V('I')},A{k}),\"-\")")
    for c, w in zip("ABCDE", (14, 9, 10, 12, 18)):
        wp.column_dimensions[c].width = w

    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    books, verses = load()
    problems = check(books, verses)
    print(f"구절 {len(verses)}편 (구약 {sum(v['track'] == '구약' for v in verses)} / 신약 {sum(v['track'] == '신약' for v in verses)})")
    for p in problems:
        print("확인 필요:", p)
    if problems:
        sys.exit(1)
    print("저장:", build(books, verses))

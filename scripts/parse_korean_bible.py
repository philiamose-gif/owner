"""개역한글·개역개정 PDF 텍스트(pdftotext)를 구절 단위 TSV로 변환.

출력 열: 약어, 장, 절, 본문, 확인필요
- 확인필요 'break': 줄이 바뀌는 곳의 띄어쓰기를 확실히 정하지 못한 구절. PDF에는 줄 끝 띄어쓰기
  정보가 없어서 성경 전체에서 같은 표현이 쓰인 방식으로 추정했으므로, 사용 전에 원본 PDF와 대조.
- 확인필요 'merged:N': 원본 PDF에서 N절과 합쳐 인쇄된 절 (예: 신6:18-19).

사용: python3 -I parse_korean_bible.py 입력.txt [입력2.txt ...] 출력.tsv
"""
import re
import sys

ABBR = ["창", "출", "레", "민", "신", "수", "삿", "룻", "삼상", "삼하", "왕상", "왕하", "대상", "대하",
        "스", "느", "에", "욥", "시", "잠", "전", "아", "사", "렘", "애", "겔", "단", "호", "욜", "암",
        "옵", "욘", "미", "나", "합", "습", "학", "슥", "말", "마", "막", "눅", "요", "행", "롬", "고전",
        "고후", "갈", "엡", "빌", "골", "살전", "살후", "딤전", "딤후", "딛", "몬", "히", "약", "벧전",
        "벧후", "요일", "요이", "요삼", "유", "계"]
ALT = "|".join(sorted(ABBR, key=len, reverse=True))
MARK = re.compile(rf"({ALT})(\d{{1,3}}):(\d{{1,3}})(?:-(\d{{1,3}}))?\s?")
HEADING = re.compile(r"<[^<>]*>\s*")
NAMES = ("창세기 출애굽기 레위기 민수기 신명기 여호수아 사사기 룻기 사무엘상 사무엘하 열왕기상 열왕기하 역대상 역대하 "
         "에스라 느헤미야 에스더 욥기 시편 잠언 전도서 아가 이사야 예레미야 예레미야애가 에스겔 다니엘 호세아 요엘 아모스 "
         "오바댜 요나 미가 나훔 하박국 스바냐 학개 스가랴 말라기 마태복음 마가복음 누가복음 요한복음 사도행전 로마서 "
         "고린도전서 고린도후서 갈라디아서 에베소서 빌립보서 골로새서 데살로니가전서 데살로니가후서 디모데전서 디모데후서 "
         "디도서 빌레몬서 히브리서 야고보서 베드로전서 베드로후서 요한일서 요한이서 요한삼서 유다서 요한계시록").split()
SKIP = re.compile(rf"^\s*(\d+|구약성경|신약성경|개역 ?한글 ?성경|개역개정.*|{'|'.join(NAMES)})\s*$")


def read_lines(paths):
    for path in paths:
        for raw in open(path, encoding="utf-8"):
            line = raw.replace("\f", "").strip()
            if line:
                yield line


def parse(paths):
    lines = list(read_lines(paths))
    # 1) 구절별 조각(줄 단위) 모으기
    verses, cur = [], None
    for line in lines:
        pos = 0
        hits = list(MARK.finditer(line))
        if not hits:
            if cur is not None and not SKIP.match(line):
                cur["parts"].append(line)
            continue
        if hits[0].start() > 0 and cur is not None:
            cur["parts"].append(line[:hits[0].start()].strip())
        for k, m in enumerate(hits):
            end = hits[k + 1].start() if k + 1 < len(hits) else len(line)
            cur = dict(abbr=m.group(1), c=int(m.group(2)), v=int(m.group(3)),
                       to=int(m.group(4)) if m.group(4) else None,
                       parts=[line[m.end():end].strip()])
            verses.append(cur)
            pos = end
    # 2) 줄 안 문장(줄바꿈과 무관한 부분)을 말뭉치로 삼아 띄어쓰기 추정
    corpus = "\n".join(HEADING.sub("", p) for vs in verses for p in vs["parts"])

    def glue_or_space(a, b):
        """(이어 쓸 문자, 확실한지). 같은 표현이 줄 안에서 3배 이상 한쪽으로 쓰였으면 확실한 것으로 본다."""
        for level, (x, y) in enumerate(((a, b), (a[-2:], b[:2]), (a[-1:], b[:2]), (a[-1:], b[:1]))):
            g, sp = corpus.count(x + y), corpus.count(x + " " + y)
            if g != sp:
                sure = level == 0 and max(g, sp) >= 3 * min(g, sp) + 2  # 낱말 전체로 판단한 경우만 확실
                return ("" if g > sp else " "), sure
        return " ", False
    # 3) 줄바꿈 이어 붙이기: 붙인 단어가 더 흔하면 붙이고, 아니면 띄움
    out = []
    for vs in verses:
        parts = [HEADING.sub("", p).strip() for p in vs["parts"] if p.strip()]
        text = parts[0] if parts else ""
        unsure = False
        for nxt in parts[1:]:
            a = text.split()[-1] if text.split() else ""
            sep, sure = glue_or_space(a, nxt.split()[0])
            unsure |= not sure
            text = text + sep + nxt
        flag = "break" if unsure else ""
        text = re.sub(r"\s+", " ", text).strip()
        out.append([vs["abbr"], vs["c"], vs["v"], text, flag])
        if vs["to"]:
            for v in range(vs["v"] + 1, vs["to"] + 1):
                out.append([vs["abbr"], vs["c"], v, text, f"merged:{vs['v']}"])
    return out


if __name__ == "__main__":
    *src, dst = sys.argv[1:]
    rows = parse(src)
    with open(dst, "w", encoding="utf-8") as fp:
        for r in rows:
            fp.write("\t".join(map(str, r)) + "\n")
    print(len(rows), "절 →", dst)

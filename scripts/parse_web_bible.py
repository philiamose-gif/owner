"""World English Bible PDF(2단 편집)를 구절 단위 TSV(약어, 장, 절, 본문)로 변환.

장 번호는 큰 글자, 절 번호는 작은 글자(9pt)로 인쇄된 점을 이용한다. 쪽 머리글(예: Genesis1:1 ... 1:31)로
장 범위를 확인하고, 각주(작은 글자 문장)는 버린다.

사용: python3 -I parse_web_bible.py web.pdf 출력.tsv
"""
import re
import sys

import pdfplumber

ABBR = ["창", "출", "레", "민", "신", "수", "삿", "룻", "삼상", "삼하", "왕상", "왕하", "대상", "대하",
        "스", "느", "에", "욥", "시", "잠", "전", "아", "사", "렘", "애", "겔", "단", "호", "욜", "암",
        "옵", "욘", "미", "나", "합", "습", "학", "슥", "말", "마", "막", "눅", "요", "행", "롬", "고전",
        "고후", "갈", "엡", "빌", "골", "살전", "살후", "딤전", "딤후", "딛", "몬", "히", "약", "벧전",
        "벧후", "요일", "요이", "요삼", "유", "계"]
ENG = ["Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy", "Joshua", "Judges", "Ruth", "1 Samuel",
       "2 Samuel", "1 Kings", "2 Kings", "1 Chronicles", "2 Chronicles", "Ezra", "Nehemiah", "Esther", "Job",
       "Psalms", "Proverbs", "Ecclesiastes", "Song of Solomon", "Isaiah", "Jeremiah", "Lamentations", "Ezekiel",
       "Daniel", "Hosea", "Joel", "Amos", "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah",
       "Haggai", "Zechariah", "Malachi", "Matthew", "Mark", "Luke", "John", "Acts", "Romans", "1 Corinthians",
       "2 Corinthians", "Galatians", "Ephesians", "Philippians", "Colossians", "1 Thessalonians",
       "2 Thessalonians", "1 Timothy", "2 Timothy", "Titus", "Philemon", "Hebrews", "James", "1 Peter",
       "2 Peter", "1 John", "2 John", "3 John", "Jude", "Revelation"]


def norm(s):
    return re.sub(r"\s+", "", s).lower()


ENG_KEY = {norm(e): i for i, e in enumerate(ENG)}


def page_book(words):
    """쪽 머리글에서 책 번호를 찾는다."""
    head = norm("".join(w["text"] for w in words if w["top"] < 90))
    m = re.match(r"([a-z0-9]+?)\d+(?::\d+)?", head)  # 한 장짜리 책은 머리글에 장 번호가 없음
    if m and m.group(1) in ENG_KEY:
        return ENG_KEY[m.group(1)]
    return None


def parse(path):
    verses = {}
    book, ch, v = -1, 0, 0
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["size"], x_tolerance=1.2, use_text_flow=True)
            pb = page_book(words)
            if pb is None:
                continue
            if pb != book:
                book, ch, v = pb, 0, 0
            body = [w for w in words if w["top"] >= 90]
            # 각주: 작은 글자로 된 숫자가 아닌 단어가 처음 나온 줄부터 아래는 버림
            notes = [w["top"] for w in body if w["size"] < 10.5 and not w["text"].isdigit()]
            cut = min(notes) - 2 if notes else 1e9
            for w in body:
                t = w["text"]
                if w["top"] >= cut and w["size"] < 10.5:
                    continue
                if t.isdigit() and w["size"] > 13:
                    ch, v = int(t), 1
                    verses.setdefault((book, ch, v), [])
                    continue
                if t.isdigit() and w["size"] < 10.5:
                    n = int(t)
                    if ch == 0:  # 장 번호 없이 시작하는 1장
                        ch = 1
                    v = n
                    verses.setdefault((book, ch, v), [])
                    continue
                if w["size"] > 13 or ch == 0 or (book, ch, v) not in verses:
                    continue
                verses[(book, ch, v)].append(t.replace("*", "").replace("†", "").replace("‡", ""))
    out = []
    for (b, c, n), toks in sorted(verses.items()):
        toks = [t for t in toks if t]
        out.append((ABBR[b], c, n, re.sub(r"\s+", " ", " ".join(toks)).strip()))
    return out


if __name__ == "__main__":
    rows = parse(sys.argv[1])
    with open(sys.argv[2], "w", encoding="utf-8") as fp:
        for r in rows:
            fp.write("\t".join(map(str, r)) + "\n")
    print(len(rows), "절 →", sys.argv[2])

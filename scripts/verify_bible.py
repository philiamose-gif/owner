"""역본 대조 검수: 프로젝트1·2의 본문 범위와 구절이 실제 성경과 맞는지 확인한다.

기준 자료 (reference/):
- krv.tsv  개역한글 (공개, 화면·음성 인용 본문)
- web.tsv  World English Bible (공개, 영어 확장용)
- private/nkrv.tsv  개역개정 (저작권 보호 — 저장소에 올리지 않음. 검수·혼입 확인용으로만 사용)

실행: python3 scripts/verify_bible.py
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")


def load(name):
    path = os.path.join(REF, name)
    if not os.path.exists(path):
        return None
    d = {}
    for line in open(path, encoding="utf-8"):
        p = line.rstrip("\n").split("\t")
        d[(p[0], int(p[1]), int(p[2]))] = (p[3], p[4] if len(p) > 4 else "")
    return d


KRV, WEB, NKRV = load("krv.tsv"), load("web.tsv"), load("private/nkrv.tsv")


def last_verse(abbr, ch, src=None):
    src = src or KRV
    vs = [v for (a, c, v) in src if a == abbr and c == ch]
    return max(vs) if vs else 0


def bounds(rng):
    """'1:26-2:3' → ((1,26),(2,3)), '3:17-18' → ((3,17),(3,18)), '23'(시편) → ((23,1),(23,끝))"""
    a, _, b = rng.partition("-")
    c1, v1 = map(int, a.split(":"))
    if not b:
        return (c1, v1), (c1, v1)
    if ":" in b:
        c2, v2 = map(int, b.split(":"))
    else:
        c2, v2 = c1, int(b)
    return (c1, v1), (c2, v2)


def text(abbr, rng, src=None):
    """구절(또는 짧은 범위) 본문과 확인 표시."""
    src = src or KRV
    (c1, v1), (c2, v2) = bounds(rng)
    out, flags = [], set()
    for c in range(c1, c2 + 1):
        lo = v1 if c == c1 else 1
        hi = v2 if c == c2 else last_verse(abbr, c, src)
        if src is KRV and (abbr, c, hi + 1) in KRV_SOURCE_MERGED:
            hi += 1
        for v in range(lo, hi + 1):
            if src is KRV and (abbr, c, v) in KRV_SOURCE_MERGED:
                flags.add("원본에서 앞 절과 붙어 인쇄됨")
                continue
            if (abbr, c, v) not in src:
                return None, {"없는 절"}
            t, fl = src[(abbr, c, v)]
            if fl.startswith("merged"):
                flags.add("앞 절과 합쳐 인쇄됨")
                continue
            if fl == "break":
                flags.add("줄바꿈 띄어쓰기 확인")
            out.append(t)
    return " ".join(out), flags


# 개역한글 PDF 원본에서 앞 절에 붙어 인쇄되어 따로 떼어지지 않은 절 (개역개정으로 존재 확인)
KRV_SOURCE_MERGED = {("삼상", 30, 31), ("시", 72, 20)}


def exists(abbr, c, v, src=None):
    if src is None and (abbr, c, v) in KRV_SOURCE_MERGED:
        return NKRV is None or (abbr, c, v) in NKRV
    return (abbr, c, v) in (src or KRV)


def check_project1(rows, psalms):
    """본문 범위·핵심 구절 존재 여부, 연속 분할 책의 빠진 절·겹친 절."""
    import build_xlsx as bx
    problems = []
    for r in rows:
        ab = r["book"]["abbr"]
        if r["book"]["name"] == "시편":
            ch = int(r["rng"])
            if not last_verse(ab, ch):
                problems.append(f"{r['code']} 시편 {ch}편 없음")
        else:
            (c1, v1), (c2, v2) = bounds(r["rng"])
            for c, v in ((c1, v1), (c2, v2)):
                if not exists(ab, c, v):
                    problems.append(f"{r['code']} {r['disp']}: {c}:{v}절 없음 (개역한글)")
        kc, kv = map(int, r["kv"].split(" ", 1)[1].split("-")[0].split(":"))
        if not exists(ab, kc, kv):
            problems.append(f"{r['code']} 핵심 구절 {r['kv']} 없음")
    # 연속 분할 책: 앞 편 끝 다음 절 = 다음 편 시작 절
    for bk in {id(r["book"]): r["book"] for r in rows}.values():
        if bk["name"] in bx.SELECTIVE or bk["name"] == "시편":
            continue
        mine = [r for r in rows if r["book"] is bk]
        ab = bk["abbr"]
        first = bounds(mine[0]["rng"])[0]
        if first != (1, 1):
            problems.append(f"{bk['name']} 첫 편이 1:1에서 시작하지 않음 ({mine[0]['rng']})")
        for cur, nxt in zip(mine, mine[1:]):
            (ec, ev) = bounds(cur["rng"])[1]
            (sc, sv) = bounds(nxt["rng"])[0]
            if (sc, sv) <= (ec, ev):
                if bounds(nxt["rng"])[1] <= (ec, ev):
                    continue  # 앞 편 안에 포함된 특별 편(예: 야베스의 기도)
                problems.append(f"{bk['name']} 겹침: {cur['rng']} / {nxt['rng']}")
                continue
            lv = last_verse(ab, ec) + (1 if (ab, ec, last_verse(ab, ec) + 1) in KRV_SOURCE_MERGED else 0)
            expect = (ec, ev + 1) if ev < lv else (ec + 1, 1)
            if (sc, sv) != expect:
                problems.append(f"{bk['name']} 빠진 절: {cur['rng']} 다음은 {expect[0]}:{expect[1]}이어야 하는데 {nxt['rng']}")
        (ec, ev) = bounds(mine[-1]["rng"])[1]
        lc = bx.CHAPTERS[bk["name"]]
        lv = last_verse(ab, lc) + (1 if (ab, lc, last_verse(ab, lc) + 1) in KRV_SOURCE_MERGED else 0)
        if (ec, ev) != (lc, lv):
            problems.append(f"{bk['name']} 마지막 편이 {lc}:{lv}까지 가지 않음 ({mine[-1]['rng']})")
    for ps in psalms:
        kc, kv = map(int, ps["kv"].split(" ", 1)[1].split("-")[0].split(":"))
        if not exists("시", kc, kv):
            problems.append(f"시편 시리즈 {ps['no']}편 핵심 구절 {ps['kv']} 없음")
    return problems


def check_project2(verses):
    problems = []
    for v in verses:
        t, flags = text(v["book"]["abbr"], v["ref"])
        if t is None:
            problems.append(f"프로젝트2 {v['disp']} 없음 (개역한글)")
    return problems


def mixed_with_nkrv(abbr, rng, candidate):
    """붙여 넣은 본문이 개역한글과 다르고 개역개정과 같으면 혼입으로 본다."""
    if NKRV is None or not candidate:
        return False
    k, _ = text(abbr, rng)
    n, _ = text(abbr, rng, NKRV)
    norm = lambda s: "".join((s or "").split())
    return norm(candidate) != norm(k) and norm(candidate) == norm(n)


if __name__ == "__main__":
    if KRV is None:
        sys.exit("reference/krv.tsv 가 없습니다")
    sys.path.insert(0, os.path.join(ROOT, "scripts"))
    sys.path.insert(0, os.path.join(ROOT, "project2"))
    import build_xlsx as bx
    import season
    import psalms as psa
    _, _, rows = bx.load()
    season.assign(rows)
    ps = psa.load(rows)
    import build as p2
    _, verses = p2.load()
    probs = check_project1(rows, ps) + check_project2(verses)
    print(f"개역한글 {len(KRV)}절 / WEB {len(WEB) if WEB else '없음'} / 개역개정 {len(NKRV) if NKRV else '없음(검수용, 선택)'}")
    print(f"검사: 프로젝트1 {len(rows)}편 + 시편 {len(ps)}편, 프로젝트2 {len(verses)}구절")
    for p in probs:
        print("확인 필요:", p)
    print("문제 없음" if not probs else f"확인 필요 {len(probs)}건")

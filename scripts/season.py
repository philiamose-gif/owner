"""시즌 배정, 트랙 순서, 시즌 롱폼 묶음, 연결 정보 계산."""
import os

SEASON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "data", "시즌1_선정.lst")
LONG_MIN = 7    # 시즌 롱폼 1편에 들어갈 최소 A 편수 (약 8분)
LONG_MAX = 15   # 남은 편을 합칠 때 넘지 않을 최대 편수


def assign(rows):
    """data/시즌1_선정.lst 에 있으면 시즌 1, 없으면 시즌 2."""
    index = {(r["book"]["name"], r["rng"]): r for r in rows}
    picked, unknown = [], []
    for raw in open(SEASON_FILE, encoding="utf-8"):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        key = tuple(line.split("|"))
        if key in index:
            picked.append(index[key])
        else:
            unknown.append(line)
    if unknown:
        raise SystemExit(f"시즌1 목록에 제작표에 없는 본문이 있습니다: {unknown}")
    dup = len(picked) - len({id(r) for r in picked})
    if dup:
        raise SystemExit(f"시즌1 목록에 중복된 본문이 {dup}개 있습니다")
    chosen = {id(r) for r in picked}
    for r in rows:
        r["season"] = 1 if id(r) in chosen else 2
    # 트랙(구약/신약)별 업로드 순서: 시즌 1을 성경 순서대로, 이어서 시즌 2를 성경 순서대로
    for t in ("구약", "신약"):
        n = 0
        for s in (1, 2):
            k = 0
            for r in rows:
                if r["book"]["testament"] == t and r["season"] == s:
                    n += 1
                    k += 1
                    r["track_no"] = n
                    r["season_no"] = k
    return rows


def long_groups(rows, bs):
    """시즌·트랙별로 통합본 B 경계를 따라 A를 7편 이상씩 묶는다."""
    groups = []
    for s in (1, 2):
        for t in ("구약", "신약"):
            cur, cur_bs = [], []
            track = []
            for b in bs:
                mine = [r for r in rows if r["b"] is b and r["season"] == s
                        and r["book"]["testament"] == t]
                if not mine:
                    continue
                cur += mine
                cur_bs.append(b)
                if len(cur) >= LONG_MIN:
                    track.append((cur, cur_bs))
                    cur, cur_bs = [], []
            if cur:
                if track and len(track[-1][0]) + len(cur) <= LONG_MAX:
                    prev, prev_bs = track.pop()
                    track.append((prev + cur, prev_bs + cur_bs))
                else:
                    track.append((cur, cur_bs))
            for members, members_bs in track:
                # 너무 긴 묶음(15편 초과)은 둘로 나눈다
                parts = [members] if len(members) <= LONG_MAX else \
                    [members[:len(members) // 2], members[len(members) // 2:]]
                for part in parts:
                    part_bs = [b for b in members_bs if any(r["b"] is b for r in part)]
                    groups.append(dict(season=s, track=t, rows=part, bs=part_bs))
    for i, g in enumerate(groups, 1):
        g["code"] = f"S{g['season']}-L{sum(1 for x in groups[:i] if x['season'] == g['season']):02d}"
        for r in g["rows"]:
            r["long"] = g["code"]
    return groups


def bridges(rows):
    """시즌1 편: 앞에서 건너뛴 시즌2 편 목록 / 시즌2 편: 시즌1의 어느 편 뒤 이야기인지."""
    for t in ("구약", "신약"):
        track = [r for r in rows if r["book"]["testament"] == t]
        last_s1 = None
        skipped = []
        for r in track:
            if r["season"] == 1:
                if skipped and last_s1 is not None:
                    names = ", ".join(f"{x['disp']} {x['title']}" for x in skipped[:3])
                    more = f" 외 {len(skipped) - 3}편" if len(skipped) > 3 else ""
                    r["bridge"] = f"앞에서 건너뛴 {len(skipped)}편(시즌2): {names}{more} → 연결 내레이션 1~2문장 필요"
                elif skipped:
                    r["bridge"] = f"앞선 {len(skipped)}편은 시즌2에서 다룸 → 시작 내레이션에 배경 한 줄"
                else:
                    r["bridge"] = "앞 편과 바로 이어짐"
                r["gap"] = len(skipped)
                last_s1, skipped = r, []
            else:
                skipped.append(r)
                if last_s1 is None:
                    r["bridge"] = f"시즌1 {t} #001 이전 이야기"
                else:
                    r["bridge"] = f"시즌1 {t} #{last_s1['track_no']:03d} '{last_s1['title']}' 다음 이야기"
                r["gap"] = 0
    return rows

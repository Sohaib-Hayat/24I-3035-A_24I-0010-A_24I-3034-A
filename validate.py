#!/usr/bin/env python3
# LIFELINE validator
# usage: python validator/validate.py cases.xlsx [--partial]
# recomputes everything from the sheets and does not trust any derived field
# --partial checks single cases only (used for the broken copies in validator/tests)
import argparse
import itertools
import re
import sys
import zipfile
from bisect import bisect_left, bisect_right
from collections import Counter, deque
from fractions import Fraction
import heapq

import numpy as np
from openpyxl import load_workbook

ALPHA = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-")
INF = float("inf")

SHEETS = [
    ("cases", ["case_id", "group_seed", "case_no", "tier", "scenario", "capacity", "target_0", "target_1", "target_2", "edit_bound", "tour_length", "sensor_nodes", "sensor_k", "safety_vars", "rations_target"], "siisiiiiiiiiiii"),
    ("crates", ["case_id", "crate_id", "cat", "weight", "value", "expiry", "serial"], "siiiiii"),
    ("locations", ["case_id", "loc_id", "type", "penalty"], "sisi"),
    ("roads", ["case_id", "road_no", "u", "v", "minutes", "cable"], "siiiii"),
    ("waves", ["case_id", "wave_no", "road_no"], "sii"),
    ("radio_log", ["case_id", "part", "text"], "sis"),
    ("radio_patterns", ["case_id", "pattern_no", "pattern", "count"], "sisi"),
    ("radio_reports", ["case_id", "report_a", "report_b"], "sss"),
    ("windows", ["case_id", "window_id", "start", "finish", "value"], "siiii"),
    ("incidents", ["case_id", "alarm_id", "tick", "priority"], "siii"),
    ("sites", ["case_id", "site_no", "x", "y"], "siii"),
    ("tour", ["case_id", "step", "site_no"], "sii"),
    ("pipeline_dims", ["case_id", "idx", "dim"], "sii"),
    ("drum_prices", ["case_id", "length", "price"], "sii"),
    ("sensor_edges", ["case_id", "u", "v"], "sii"),
    ("sensor_cover", ["case_id", "node"], "si"),
    ("safety_rules", ["case_id", "rule_no", "lit_1", "lit_2", "lit_3"], "siiii"),
    ("safety_planted", ["case_id", "switch", "value"], "sii"),
    ("rations", ["case_id", "item", "value", "planted"], "siii"),
    ("claims", ["case_id", "cargo_opt", "window_count_opt", "window_value_opt", "drum_opt", "tour_opt", "cover_min"], "siiiiii"),
]

RANGES = {
    "S": dict(crates=(8, 20), V=(6, 12), log=(200, 500), rep=(10, 30), windows=(8, 20), alarms=(10, 40), sites=(5, 9), stages=(4, 8), drum=(5, 14), nodes=(6, 14), vars=(6, 15), rations=(8, 20)),
    "M": dict(crates=(40, 200), V=(30, 150), log=(2000, 10000), rep=(30, 150), windows=(50, 300), alarms=(200, 1000), sites=(10, 16), stages=(15, 60), drum=(40, 200), nodes=(20, 40), vars=(20, 50), rations=(24, 40)),
    "L": dict(crates=(500, 1500), V=(500, 2000), log=(20000, 100000), rep=(200, 1000), windows=(1000, 3000), alarms=(2000, 5000), sites=(200, 500), stages=(100, 300), drum=(500, 2000), nodes=(100, 400), vars=(20, 50), rations=(50, 100)),
}


def tier_of(k):
    if k < 20:
        return "S"
    if k < 40:
        return "M"
    if k < 46:
        return "L"
    return "X"


def rng_of(tier, key, k):
    t = "L" if tier == "X" else tier
    lo, hi = RANGES[t][key]
    if key == "crates" and k == 46:
        return 1500, 3000
    if key == "crates" and k == 47:
        return 12, 60
    if key == "V" and k == 48:
        return 1500, 3000
    if key == "log" and k == 49:
        return 50000, 100000
    return lo, hi


# loading and structure checks

def load(path):
    wb = load_workbook(path, read_only=True)
    names = list(wb.sheetnames)
    data = {}
    for name in names:
        data[name] = [tuple(r) for r in wb[name].iter_rows(values_only=True)]
    merged = 0
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                merged += z.read(n).count(b"<mergeCell ")
    return names, data, merged


class Report:
    def __init__(self):
        self.glob = []
        self.cases = {}

    def g(self, code, msg):
        self.glob.append((code, msg))

    def c(self, cid, code, msg):
        self.cases.setdefault(cid, []).append((code, msg))


def type_ok(v, t):
    if t == "s":
        return isinstance(v, str)
    return isinstance(v, int) and not isinstance(v, bool)


def structure(names, data, merged, rep):
    exp = [s[0] for s in SHEETS]
    if names != exp:
        rep.g("V1", "sheets are %s, expected exactly %s in this order" % (names, exp))
    if merged:
        rep.g("V1", "merged cells found")
    grouped = {}
    for name, hdr, types in SHEETS:
        grouped[name] = {}
        if name not in data:
            continue
        rows = data[name]
        if not rows or list(rows[0]) != hdr:
            rep.g("V1", "sheet %s has wrong header %s" % (name, list(rows[0]) if rows else None))
            continue
        prev = None
        for r in rows[1:]:
            if all(v is None for v in r):
                rep.g("V1", "sheet %s has an empty row" % name)
                continue
            if len(r) != len(hdr) or any(v is None for v in r):
                rep.g("V1", "sheet %s has a row with missing cells %s" % (name, r[:3]))
                continue
            bad = False
            for v, t in zip(r, types):
                if not type_ok(v, t):
                    rep.g("V1", "sheet %s has a wrongly typed cell %r in row %s" % (name, v, r[:2]))
                    bad = True
                    break
                if t == "s":
                    if not v.isascii() or v.startswith("="):
                        rep.g("V1", "sheet %s has a non-ASCII text or formula cell in row %s" % (name, r[:2]))
                        bad = True
                        break
            if bad:
                continue
            key = (r[0], r[1]) if types[1] == "i" else (r[0],)
            if prev is not None and key < prev:
                rep.g("V1", "sheet %s is not sorted by case_id then second column near %s" % (name, r[:2]))
            prev = key
            grouped[name].setdefault(r[0], []).append(r)
    return grouped


# helpers (algorithms are different from the generator on purpose)

def bfs(V, adj, src, closed=None):
    dist = [-1] * V
    dist[src] = 0
    q = deque([src])
    while q:
        u = q.popleft()
        for v, m, ei in adj[u]:
            if closed is not None and ei in closed:
                continue
            if dist[v] < 0:
                dist[v] = dist[u] + 1
                q.append(v)
    return dist


def dijkstra(V, adj, pen):
    dist = [INF] * V
    dist[0] = 0
    pq = [(0, 0)]
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist[u]:
            continue
        for v, m, ei in adj[u]:
            nd = d + m + pen[v]
            if nd < dist[v]:
                dist[v] = nd
                heapq.heappush(pq, (nd, v))
    return dist


def min_hop_cost(V, adj, pen):
    # cheapest cost among routes that use the fewest roads
    lvl = bfs(V, adj, 0)
    order = sorted(range(V), key=lambda x: lvl[x])
    g = [INF] * V
    g[0] = 0
    for u in order:
        if g[u] == INF:
            continue
        for v, m, ei in adj[u]:
            if lvl[v] == lvl[u] + 1:
                g[v] = min(g[v], g[u] + m + pen[v])
    return g


def route_trap_targets(V, adj, pen, targets):
    g = min_hop_cost(V, adj, pen)
    d = dijkstra(V, adj, pen)
    return [t for t in targets if g[t] > d[t]], [t for t in targets if 2 * g[t] >= 3 * d[t]]


def edit_dist(a, b):
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def count_all(text, pat):
    c = 0
    i = text.find(pat)
    while i >= 0:
        c += 1
        i = text.find(pat, i + 1)
    return c


def knap_dp(ws, vs, cap):
    best = [0] * (cap + 1)
    for w, v in zip(ws, vs):
        for j in range(cap, w - 1, -1):
            if best[j - w] + v > best[j]:
                best[j] = best[j - w] + v
    return best[cap]


def greedy_vpk(crates, cap):
    order = sorted(crates, key=lambda c: (-Fraction(c["value"], c["weight"]), c["id"]))
    left = cap
    tot = 0
    for c in order:
        if c["weight"] <= left:
            left -= c["weight"]
            tot += c["value"]
    return tot


def brute_knap(ws, vs, cap):
    W = np.zeros(1, dtype=np.int64)
    Vv = np.zeros(1, dtype=np.int64)
    for w, v in zip(ws, vs):
        W = np.concatenate([W, W + w])
        Vv = np.concatenate([Vv, Vv + v])
    return int(Vv[W <= cap].max())


def brute_windows(wins):
    n = len(wins)
    conflict = [0] * n
    for i in range(n):
        for j in range(n):
            if i != j and wins[i][0] < wins[j][1] and wins[j][0] < wins[i][1]:
                conflict[i] |= 1 << j
    best = [0, 0]
    remv = [0] * (n + 1)
    for i in range(n - 1, -1, -1):
        remv[i] = remv[i + 1] + wins[i][2]

    def rec(i, blocked, cnt, val):
        if cnt + (n - i) <= best[0] and val + remv[i] <= best[1]:
            return
        if i == n:
            best[0] = max(best[0], cnt)
            best[1] = max(best[1], val)
            return
        if not (blocked >> i) & 1:
            rec(i + 1, blocked | conflict[i], cnt + 1, val + wins[i][2])
        rec(i + 1, blocked, cnt, val)

    rec(0, 0, 0, 0)
    return best[0], best[1]


def brute_tour(pts):
    n = len(pts)
    dm = [[abs(a[0] - b[0]) + abs(a[1] - b[1]) for b in pts] for a in pts]
    best = INF
    for perm in itertools.permutations(range(1, n)):
        t = dm[0][perm[0]] + dm[perm[-1]][0]
        for i in range(len(perm) - 1):
            t += dm[perm[i]][perm[i + 1]]
        if t < best:
            best = t
    return best


def brute_drum(prices):
    def rec(rem):
        if rem == 0:
            return 0
        return max(prices[i - 1] + rec(rem - i) for i in range(1, rem + 1))
    return rec(len(prices))


def brute_cover(n, edges):
    for size in range(n + 1):
        for comb in itertools.combinations(range(n), size):
            s = set(comb)
            if all(u in s or v in s for u, v in edges):
                return size
    return n


def nn_len(pts):
    n = len(pts)
    seen = [False] * n
    seen[0] = True
    cur = 0
    tot = 0
    for _ in range(n - 1):
        best, bd = -1, INF
        for j in range(n):
            if not seen[j]:
                d = abs(pts[cur][0] - pts[j][0]) + abs(pts[cur][1] - pts[j][1])
                if d < bd:
                    best, bd = j, d
        seen[best] = True
        tot += bd
        cur = best
    return tot + abs(pts[cur][0] - pts[0][0]) + abs(pts[cur][1] - pts[0][1])


def mean_overlap_x(wins):
    starts = sorted(w[0] for w in wins)
    fins = sorted(w[1] for w in wins)
    tot = 0
    for s, f, v in wins:
        tot += bisect_left(starts, f) - bisect_right(fins, s) - 1
    return Fraction(tot, len(wins))


def monotone(seq):
    up = all(seq[i] <= seq[i + 1] for i in range(len(seq) - 1))
    down = all(seq[i] >= seq[i + 1] for i in range(len(seq) - 1))
    return up or down


# per case checks

def check_case(cid, C, rep, info, partial):
    def err(code, msg):
        rep.c(cid, code, msg)

    cr = C["cases"]
    if len(cr) != 1:
        err("V1", "cases sheet must hold exactly one row for this case")
        return
    (_, s, k, tier, sc, cap, t0, t1, t2, edit_bound, tour_length, sn, sk, svars, rtarget) = cr[0]
    if not re.fullmatch(r"\d+-\d\d", cid) or cid != "%d-%02d" % (s, k):
        err("V1", "case_id %s does not match group_seed %s and case_no %s" % (cid, s, k))
    if not 0 <= k <= 49:
        err("V1", "case_no %s outside 0..49" % k)
        return
    if tier != tier_of(k):
        err("V1", "tier %s is wrong for case_no %d" % (tier, k))
        tier = tier_of(k)
    if sc != (k % 6) + 1:
        err("V1", "scenario %s is wrong for case_no %d" % (sc, k))
    scn = (k % 6) + 1
    targets = [t0, t1, t2]

    # crates
    cs = C["crates"]
    crates = [dict(id=r[1], cat=r[2], weight=r[3], value=r[4], expiry=r[5], serial=r[6]) for r in cs]
    n = len(crates)
    lo, hi = rng_of(tier, "crates", k)
    if not lo <= n <= hi:
        err("V2", "crate count %d outside %d..%d" % (n, lo, hi))
    if [c["id"] for c in crates] != list(range(1, n + 1)):
        err("V3", "crate_id is not 1..n")
    for c in crates:
        if not (0 <= c["cat"] <= 4 and 1 <= c["weight"] <= 20 and 10 <= c["value"] <= 100 and 0 <= c["expiry"] <= 999 and 10_000_000 <= c["serial"] <= 99_999_999):
            err("V2", "crate %s has a field outside its range" % c["id"])
            break
    if len(set(c["serial"] for c in crates)) != n:
        err("V3", "crate serials are not distinct")
    total_w = sum(c["weight"] for c in crates)
    if not (cap >= 1 and 10 * cap >= total_w and 10 * cap <= 7 * total_w):
        err("V2", "capacity %d is not 10%% to 70%% of total weight %d" % (cap, total_w))

    # locations and roads
    locs = C["locations"]
    V = len(locs)
    lo, hi = rng_of(tier, "V", k)
    if not lo <= V <= hi:
        err("V2", "location count %d outside %d..%d" % (V, lo, hi))
    if [r[1] for r in locs] != list(range(V)):
        err("V3", "loc_id is not 0..V-1")
    ltype = {r[1]: r[2] for r in locs}
    pen = [r[3] for r in sorted(locs, key=lambda r: r[1])]
    if any(r[2] not in ("D", "S", "H", "J") or r[3] not in (0, 1, 2, 4, 6) for r in locs):
        err("V2", "location type or penalty is outside its set")
    if ltype.get(0) != "D" or (pen and pen[0] != 0):
        err("V2", "location 0 must be the HQ: type D, penalty 0")
    if not any(t == "S" for t in ltype.values()) or not any(t == "H" for t in ltype.values()):
        err("V2", "need at least one shelter and one hospital")
    for t in targets:
        if t not in ltype or ltype[t] not in ("S", "H") or t == 0:
            err("V3", "target %s is not an existing shelter or hospital other than the HQ" % t)
    if len(set(targets)) != 3:
        err("V2", "targets are not distinct")
    rd = C["roads"]
    E = len(rd)
    if [r[1] for r in rd] != list(range(E)):
        err("V3", "road_no is not 0..E-1")
    if not V - 1 <= E <= 3 * V:
        err("V2", "road count %d outside V-1..3V" % E)
    seen = set()
    adj = [[] for _ in range(V)]
    roads_ok = True
    for r in rd:
        _, rno, u, v, m, cab = r
        if not (0 <= u < v < V):
            err("V3", "road %d has endpoints that do not exist or are not u<v" % rno)
            roads_ok = False
            continue
        if (u, v) in seen:
            err("V2", "road %d repeats a pair" % rno)
        seen.add((u, v))
        if not (1 <= m <= 60 and 1 <= cab <= 100):
            err("V2", "road %d minutes or cable outside range" % rno)
        adj[u].append((v, m, rno))
        adj[v].append((u, m, rno))
    dist0 = None
    if roads_ok and V > 0:
        dist0 = bfs(V, adj, 0)
        if min(dist0) < 0:
            err("V4", "network is not connected before closures")
    wv = C["waves"]
    wnos = sorted(set(r[1] for r in wv))
    if wnos != list(range(1, len(wnos) + 1)):
        err("V3", "wave_no values are not 1..w")
    if not 2 <= len(wnos) <= 5:
        err("V2", "wave count %d outside 2..5" % len(wnos))
    wroads = [r[2] for r in wv]
    if any(not 0 <= x < E for x in wroads):
        err("V3", "a wave refers to a road that does not exist")
    if len(set(wroads)) != len(wroads):
        err("V4", "waves are not disjoint")
    for w in wnos:
        sz = sum(1 for r in wv if r[1] == w)
        if not 1 <= sz <= max(1, E // 4):
            err("V4", "wave %d has %d roads, limit is 1..%d" % (w, sz, max(1, E // 4)))

    # radio
    lg = C["radio_log"]
    if [r[1] for r in lg] != list(range(1, len(lg) + 1)):
        err("V5", "log parts are not numbered 1,2,...")
    lg = sorted(lg, key=lambda r: r[1])
    text = "".join(r[2] for r in lg)
    if any(len(r[2]) == 0 or len(r[2]) > 10000 for r in lg):
        err("V5", "a log part is empty or longer than 10,000 symbols")
    N = len(text)
    lo, hi = rng_of(tier, "log", k)
    if not lo <= N <= hi:
        err("V2", "log length %d outside %d..%d" % (N, lo, hi))
    if not set(text) <= ALPHA:
        err("V2", "log holds symbols outside the 38-symbol alphabet")
    pt = C["radio_patterns"]
    if [r[1] for r in pt] != list(range(6)) or [len(r[2]) for r in pt] != [2, 4, 8, 12, 20, 40]:
        err("V3", "patterns must be numbered 0..5 with lengths 2,4,8,12,20,40")
    else:
        counts = []
        for r in pt:
            if not set(r[2]) <= ALPHA:
                err("V2", "pattern %d has symbols outside the alphabet" % r[1])
            real = count_all(text, r[2])
            counts.append(real)
            if real != r[3]:
                err("V5", "pattern %d count is %d, recomputed %d" % (r[1], r[3], real))
        if sum(1 for c in counts if c > 0) < 2 or min(counts) > 0:
            err("V2", "need at least two patterns that occur and one that does not")
    rr = C["radio_reports"]
    ea = None
    if len(rr) != 1:
        err("V1", "radio_reports must hold one row")
    else:
        a, b = rr[0][1], rr[0][2]
        lo, hi = rng_of(tier, "rep", k)
        if not (lo <= len(a) <= hi and lo <= len(b) <= hi):
            err("V2", "report lengths %d, %d outside %d..%d" % (len(a), len(b), lo, hi))
        if not (set(a) <= ALPHA and set(b) <= ALPHA):
            err("V2", "reports use symbols outside the alphabet")
        ea = edit_dist(a, b)
        if not ea <= edit_bound <= len(a) + len(b):
            err("V5", "edit_bound %d is not between the true distance %d and %d" % (edit_bound, ea, len(a) + len(b)))

    # windows
    wn = C["windows"]
    wins = [(r[2], r[3], r[4]) for r in wn]
    lo, hi = rng_of(tier, "windows", k)
    if not lo <= len(wins) <= hi:
        err("V2", "window count %d outside %d..%d" % (len(wins), lo, hi))
    if [r[1] for r in wn] != list(range(1, len(wn) + 1)):
        err("V3", "window_id is not 1..n")
    for st, fi, v in wins:
        if not (0 <= st < fi <= 1440 and 5 <= fi - st <= 480 and 1 <= v <= 100):
            err("V2", "a window is outside its ranges")
            break
    if len(wins) >= 8:
        touch = sum(1 for i in range(len(wins)) for j in range(len(wins)) if i != j and wins[i][1] == wins[j][0]) if len(wins) <= 400 else None
        if touch is None:
            starts = Counter(w[0] for w in wins)
            touch = sum(starts[w[1]] for w in wins)
        if touch < 2:
            err("V2", "fewer than two touching pairs of windows")

    # incidents
    ic = C["incidents"]
    inc = [(r[2], r[3]) for r in ic]
    lo, hi = rng_of(tier, "alarms", k)
    if not lo <= len(inc) <= hi:
        err("V2", "alarm count %d outside %d..%d" % (len(inc), lo, hi))
    if [r[1] for r in ic] != list(range(1, len(ic) + 1)):
        err("V3", "alarm_id is not 1..n")
    ticks = [t for t, p in inc]
    pris = [p for t, p in inc]
    if not ticks or ticks[0] != 0 or any(ticks[i] > ticks[i + 1] for i in range(len(ticks) - 1)):
        err("V2", "ticks must start at 0 and never decrease")
    if any(not 1 <= p <= 10 for p in pris):
        err("V2", "priority outside 1..10")
    rate = Fraction(len(inc), ticks[-1] + 1) if ticks else Fraction(0)
    if tier in ("M", "L"):
        if not Fraction(12, 10) <= rate <= Fraction(18, 10):
            err("V2", "alarms per tick %s outside 1.2..1.8" % float(rate))
        if len(set(pris)) < 6 or 1 not in pris:
            err("V2", "need at least 6 distinct priorities including 1")

    # sites and tour
    st = C["sites"]
    pts = [(r[2], r[3]) for r in st]
    n_s = len(pts)
    lo, hi = rng_of(tier, "sites", k)
    if not lo <= n_s <= hi:
        err("V2", "site count %d outside %d..%d" % (n_s, lo, hi))
    if [r[1] for r in st] != list(range(n_s)):
        err("V3", "site_no is not 0..n-1")
    if any(not (0 <= x <= 999 and 0 <= y <= 999) for x, y in pts) or len(set(pts)) != n_s:
        err("V2", "site coordinates outside 0..999 or two sites equal")
    tr = C["tour"]
    order = [r[2] for r in tr]
    if [r[1] for r in tr] != list(range(len(tr))) or sorted(order) != list(range(n_s)) or order[:1] != [0]:
        err("V6", "tour is not a permutation of all sites starting at site 0")
    else:
        real = sum(abs(pts[order[i]][0] - pts[order[(i + 1) % n_s]][0]) + abs(pts[order[i]][1] - pts[order[(i + 1) % n_s]][1]) for i in range(n_s))
        if real != tour_length:
            err("V6", "tour_length %d but the closed Manhattan length is %d" % (tour_length, real))
        if tier in ("L", "X") and real > nn_len(pts):
            err("V6", "planted tour is longer than the nearest-neighbour tour")

    # pipeline and drum
    pp = C["pipeline_dims"]
    if [r[1] for r in pp] != list(range(len(pp))):
        err("V3", "pipeline idx is not 0..n")
    lo, hi = rng_of(tier, "stages", k)
    if not lo <= len(pp) - 1 <= hi:
        err("V2", "pipeline stage count %d outside %d..%d" % (len(pp) - 1, lo, hi))
    if any(not 2 <= r[2] <= 100 for r in pp):
        err("V2", "a pipeline dimension is outside 2..100")
    dr = C["drum_prices"]
    prices = [r[2] for r in dr]
    lo, hi = rng_of(tier, "drum", k)
    if [r[1] for r in dr] != list(range(1, len(dr) + 1)):
        err("V3", "drum length is not 1..L")
    if not lo <= len(dr) <= hi:
        err("V2", "drum length %d outside %d..%d" % (len(dr), lo, hi))
    if any(not 1 <= p <= 500 for p in prices):
        err("V2", "a drum price is outside 1..500")
    if prices and all(prices[j - 1] * 1 == prices[0] * j for j in range(1, len(prices) + 1)):
        err("V2", "drum prices are all proportional to length")

    # sensors
    lo, hi = rng_of(tier, "nodes", k)
    if not lo <= sn <= hi:
        err("V2", "sensor_nodes %d outside %d..%d" % (sn, lo, hi))
    se = [(r[1], r[2]) for r in C["sensor_edges"]]
    if any(not 0 <= u < v < sn for u, v in se) or len(set(se)) != len(se):
        err("V3", "sensor edges must satisfy 0<=u<v<nodes with no repeats")
    if len(se) < sn - 1:
        err("V2", "fewer than n-1 sensor edges")
    cv = [r[1] for r in C["sensor_cover"]]
    if len(set(cv)) != len(cv) or any(not 0 <= c < sn for c in cv):
        err("V6", "cover has repeated or missing nodes")
    if len(cv) != sk or not (2 <= sk and 2 * sk <= sn):
        err("V6", "cover size %d, sensor_k %d, need 2<=k<=n/2" % (len(cv), sk))
    cset = set(cv)
    if any(u not in cset and v not in cset for u, v in se):
        err("V6", "planted cover misses a sensor edge")

    # safety
    lo, hi = rng_of(tier, "vars", k)
    if not lo <= svars <= hi:
        err("V2", "safety_vars %d outside %d..%d" % (svars, lo, hi))
    sr = C["safety_rules"]
    if [r[1] for r in sr] != list(range(1, len(sr) + 1)):
        err("V3", "rule_no is not 1..m")
    for r in sr:
        lits = r[2:5]
        if any(x == 0 or abs(x) > svars for x in lits) or len(set(abs(x) for x in lits)) != 3:
            err("V2", "rule %d must use three non-zero literals on three different switches" % r[1])
            break
    ratio_ok = svars > 0 and 2 * svars <= len(sr) <= 6 * svars
    if not ratio_ok:
        err("V2", "rules per switch %s outside 2.0..6.0" % (len(sr) / svars if svars else None))
    sp = C["safety_planted"]
    if [r[1] for r in sp] != list(range(1, svars + 1)) or any(r[2] not in (0, 1) for r in sp):
        err("V3", "safety_planted must list switches 1..vars with values 0 or 1")
    else:
        val = {r[1]: r[2] for r in sp}
        for r in sr:
            if not any((x > 0) == (val[abs(x)] == 1) for x in r[2:5] if abs(x) in val):
                err("V6", "planted assignment breaks rule %d" % r[1])
                break

    # rations
    ra = C["rations"]
    lo, hi = rng_of(tier, "rations", k)
    if not lo <= len(ra) <= hi:
        err("V2", "ration quantity count %d outside %d..%d" % (len(ra), lo, hi))
    if [r[1] for r in ra] != list(range(len(ra))):
        err("V3", "ration item is not 0..q-1")
    if any(not 1 <= r[2] <= 1000 or r[3] not in (0, 1) for r in ra):
        err("V2", "ration value or planted flag outside range")
    if rtarget <= 0 or sum(r[2] for r in ra if r[3] == 1) != rtarget:
        err("V6", "planted ration values do not add up to rations_target")

    # claims
    cl = C["claims"]
    if tier == "S":
        if len(cl) != 1:
            err("V5", "tier S needs exactly one claims row")
        else:
            _, c_cargo, c_wc, c_wv, c_drum, c_tour, c_cover = cl[0]
            ws = [c["weight"] for c in crates]
            vs = [c["value"] for c in crates]
            real = brute_knap(ws, vs, cap)
            if c_cargo != real:
                err("V5", "cargo_opt %d, exhaustive search gives %d" % (c_cargo, real))
            info["greedy_fail"] = greedy_vpk(crates, cap) < real
            rwc, rwv = brute_windows(wins)
            if c_wc != rwc or c_wv != rwv:
                err("V5", "window claims %d/%d, exhaustive search gives %d/%d" % (c_wc, c_wv, rwc, rwv))
            if c_drum != brute_drum(prices):
                err("V5", "drum_opt is wrong")
            if c_tour != brute_tour(pts):
                err("V5", "tour_opt is wrong")
            if c_cover != brute_cover(sn, se):
                err("V5", "cover_min is wrong")
            if c_cover > sk:
                err("V5", "cover_min is larger than the planted cover")
    elif cl:
        err("V5", "claims rows exist for a tier that is not S")

    # route trap property on the whole network (used by variety rule 2 and trap 48)
    if roads_ok and dist0 is not None and min(dist0) >= 0 and all(t in ltype for t in targets):
        strict, ratio = route_trap_targets(V, adj, pen, targets)
        info["route_trap"] = bool(strict)
        info["route_ratio"] = bool(ratio)
    info["ratio"] = Fraction(len(sr), svars) if svars else None

    # signature rules, cases 0 to 45
    if k <= 45:
        deg2 = Fraction(2 * E, V) if V else Fraction(0)
        cats = Counter(c["cat"] for c in crates)
        nwaves = len(wnos)
        if scn == 1:
            if not Fraction(22, 10) <= deg2 <= Fraction(32, 10):
                err("V7", "delta flood: average degree %s outside 2.2..3.2" % float(deg2))
            if 5 * sum(1 for p in pen if p >= 2) < 2 * V:
                err("V7", "delta flood: fewer than 40%% of locations have penalty 2 or more")
            if nwaves < 3:
                err("V7", "delta flood: fewer than 3 waves")
            if 100 * (cats[0] + cats[1]) < 55 * n:
                err("V7", "delta flood: food and water crates below 55%")
        if scn == 2:
            if not Fraction(2) <= deg2 <= Fraction(26, 10):
                err("V7", "mountain earthquake: average degree %s outside 2.0..2.6" % float(deg2))
            if 100 * (cats[2] + cats[3]) < 45 * n:
                err("V7", "mountain earthquake: medicine and tool crates below 45%")
            if pris and sum(pris) < 6 * len(pris):
                err("V7", "mountain earthquake: mean alarm priority below 6.0")
        if scn == 3:
            if 4 * sum(1 for t in ltype.values() if t == "S") < V:
                err("V7", "coastal cyclone: fewer than 25% shelters")
            if nwaves < 3:
                err("V7", "coastal cyclone: fewer than 3 waves")
            if wins and mean_overlap_x(wins) < 2:
                err("V7", "coastal cyclone: windows overlap fewer than 2 others on average")
            if len(set(text)) < 30:
                err("V7", "coastal cyclone: log uses fewer than 30 symbols")
        if scn == 4:
            if deg2 < Fraction(7, 2):
                err("V7", "urban flash flood: average degree below 3.5")
            if tier in ("M", "L") and rate < Fraction(3, 2):
                err("V7", "urban flash flood: fewer than 1.5 alarms per tick")
            if 10 * cats[4] < n:
                err("V7", "urban flash flood: rare crates below 10%")
        if scn == 5:
            if deg2 > Fraction(12, 5):
                err("V7", "landslide corridor: average degree above 2.4")
            if any(len(a) > 4 for a in adj):
                err("V7", "landslide corridor: a location has more than 4 roads")
            if dist0 is not None and max(dist0) < V // 4:
                err("V7", "landslide corridor: farthest location is nearer than floor(V/4)")
            if nwaves < 3:
                err("V7", "landslide corridor: fewer than 3 waves")
            if roads_ok:
                after = bfs(V, adj, 0, set(wroads))
                if all(after[t] >= 0 for t in targets if 0 <= t < V):
                    err("V7", "landslide corridor: all waves together leave every target reachable")
        if scn == 6:
            if 100 * (cats[2] + cats[1]) < 55 * n:
                err("V7", "outbreak camp: medicine and water crates below 55%")
            if 2 * sum(1 for c in crates if c["expiry"] <= 300) < n:
                err("V7", "outbreak camp: fewer than half of the crates expire by hour 300")
            if N and 4 * max(Counter(text).values()) < N:
                err("V7", "outbreak camp: no symbol reaches 25% of the log")

    # trap cases
    if k == 46:
        vpk = [Fraction(c["value"], c["weight"]) for c in crates]
        flags = [monotone([c["value"] for c in crates]), monotone([c["weight"] for c in crates]), monotone([c["expiry"] for c in crates]), monotone([c["serial"] for c in crates]), monotone(vpk)]
        if sum(flags) < 3:
            err("V8", "dock-order trap: only %d of the five fields are monotone" % sum(flags))
    if k == 47:
        if crates and cap >= 1:
            opt = knap_dp([c["weight"] for c in crates], [c["value"] for c in crates], cap)
            gv = greedy_vpk(crates, cap)
            if 5 * gv > 3 * opt:
                err("V8", "best-looking-crate trap: greedy reaches %d of %d which is above 60%%" % (gv, opt))
    if k == 48:
        if info.get("route_ratio") is not True:
            err("V8", "shortest-looks-cheapest trap: no target has fewest-roads cost of at least 1.5 times the cheapest")
        if roads_ok and dist0 is not None and min(dist0) >= 0:
            far = max(range(V), key=lambda x: dist0[x])
            d2 = bfs(V, adj, far)
            if max(d2) < 500:
                far2 = max(range(V), key=lambda x: d2[x])
                diam = max(max(bfs(V, adj, u)) for u in range(V)) if V <= 3000 else max(d2)
            else:
                diam = max(d2)
            if diam < 500:
                err("V8", "shortest-looks-cheapest trap: longest distance is %d roads, need 500" % diam)
    if k == 49:
        if N and not 50000 <= N <= 100000:
            err("V8", "near-miss trap: log length outside 50,000..100,000")
        best = 0
        for b in range(1, 9):
            tot = 0
            for r in range(b):
                tot += max(Counter(text[r::b]).values()) if text[r::b] else 0
            best = max(best, tot)
        if N and 10 * best < 9 * N:
            err("V8", "near-miss trap: only %.1f%% of the log repeats one block of at most 8 symbols" % (100.0 * best / N))
        near = 0
        for r in pt:
            L = len(r[2])
            pre = r[2][:-1]
            hits = sum(1 for i in range(N - L + 1) if text.startswith(pre, i))
            if 2 * hits >= N - L + 1:
                near += 1
        if near < 4:
            err("V8", "near-miss trap: only %d patterns have their first length-1 symbols matching at half the positions" % near)

    # module fingerprints for variety rule 4
    mods = {
        "crates": (len(crates), tuple((c["cat"], c["weight"], c["value"], c["expiry"], c["serial"]) for c in crates)),
        "network": (V, tuple(tuple(r[2:]) for r in rd), tuple(tuple(r[2:]) for r in locs)),
        "radio_log": (N, text),
        "radio_reports": (len(rr), tuple(rr[0][1:]) if rr else ()),
        "windows": (len(wins), tuple(wins)),
        "incidents": (len(inc), tuple(inc)),
        "sites": (n_s, tuple(pts)),
        "pipeline_dims": (len(pp), tuple(r[2] for r in pp)),
        "drum_prices": (len(prices), tuple(prices)),
        "sensor_edges": (sn, tuple(se)),
        "safety_rules": (len(sr), tuple(tuple(r[2:]) for r in sr)),
        "rations": (len(ra), tuple(tuple(r[2:]) for r in ra)),
    }
    info["mods"] = {m: hash(v[1]) for m, v in mods.items() if v[0] >= 8}


def check_global(infos, rep):
    tierS = [cid for cid, i in infos.items() if i["tier"] == "S"]
    if sum(1 for cid in tierS if infos[cid].get("greedy_fail")) < 5:
        rep.g("V7", "variety 1: fewer than 5 tier S cases where value-per-kg greedy is not optimal")
    if sum(1 for cid in tierS if infos[cid].get("route_trap")) < 5:
        rep.g("V7", "variety 2: fewer than 5 tier S cases where fewest roads is not cheapest for some target")
    ml = [cid for cid, i in infos.items() if i["tier"] in ("M", "L")]
    if len(ml) != 26:
        rep.g("V7", "variety 3: expected 26 cases of tiers M and L, found %d" % len(ml))
    rs = [infos[c]["ratio"] for c in ml if infos[c].get("ratio") is not None]
    if rs:
        if any(not Fraction(5, 2) <= r <= 6 for r in rs):
            rep.g("V7", "variety 3: a rules-per-switch ratio on tiers M and L is outside 2.5..6.0")
        if sum(1 for r in rs if r < Fraction(7, 2)) < 5:
            rep.g("V7", "variety 3: fewer than 5 cases below 3.5")
        if sum(1 for r in rs if Fraction(7, 2) <= r <= Fraction(24, 5)) < 8:
            rep.g("V7", "variety 3: fewer than 8 cases from 3.5 to 4.8")
        if sum(1 for r in rs if r > Fraction(24, 5)) < 5:
            rep.g("V7", "variety 3: fewer than 5 cases above 4.8")
    seen = {}
    for cid, i in infos.items():
        for m, h in i.get("mods", {}).items():
            key = (m, h)
            if key in seen:
                rep.g("V7", "variety 4: cases %s and %s are equal in module %s" % (seen[key], cid, m))
            else:
                seen[key] = cid


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--partial", action="store_true", help="check single cases only, skip the 50-case rules")
    args = ap.parse_args()
    names, data, merged = load(args.path)
    rep = Report()
    grouped = structure(names, data, merged, rep)
    ids = sorted(grouped["cases"].keys()) if "cases" in grouped else []
    if not args.partial:
        seeds = set(grouped["cases"][i][0][1] for i in ids) if ids else set()
        if len(ids) != 50:
            rep.g("V1", "found %d case_ids, expected 50" % len(ids))
        if len(seeds) == 1:
            s = seeds.pop()
            want = ["%d-%02d" % (s, k) for k in range(50)]
            if ids != want:
                rep.g("V1", "case_ids are not exactly %s .. %s" % (want[0], want[-1]))
    allids = set(ids)
    for name in grouped:
        for cid in grouped[name]:
            if cid not in allids:
                rep.g("V1", "sheet %s has rows for case_id %s that is not in the cases sheet" % (name, cid))
    infos = {}
    for cid in ids:
        C = {name: grouped[name].get(cid, []) for name, _, _ in SHEETS}
        info = {"tier": C["cases"][0][3] if C["cases"] else "?"}
        infos[cid] = info
        check_case(cid, C, rep, info, args.partial)
        info["tier"] = tier_of(C["cases"][0][2]) if C["cases"] and isinstance(C["cases"][0][2], int) and 0 <= C["cases"][0][2] <= 49 else "?"
    if not args.partial and len(ids) == 50:
        check_global(infos, rep)
    ok = True
    for cid in ids:
        errs = rep.cases.get(cid, [])
        if errs:
            ok = False
            print("%s FAIL" % cid)
            for code, msg in errs:
                print("    %s %s" % (code, msg))
        else:
            print("%s PASS" % cid)
    for code, msg in rep.glob:
        ok = False
        print("GLOBAL FAIL %s %s" % (code, msg))
    print("RESULT: %s" % ("PASS" if ok else "FAIL"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

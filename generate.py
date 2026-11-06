#!/usr/bin/env python3
# LIFELINE data generator
# usage: python generator/generate.py --seed GROUP_NUMBER --out cases.xlsx
# all randomness comes from the group seed and the case number
import argparse
import heapq
import itertools
import math
import random
from bisect import bisect_left, bisect_right
from collections import deque
from fractions import Fraction

from openpyxl import Workbook

ALPHABET = [chr(c) for c in range(65, 91)] + [str(d) for d in range(10)] + ["_", "-"]
PENALTIES = [0, 1, 2, 4, 6]
INF = float("inf")

SHEETS = [
    ("cases", ["case_id", "group_seed", "case_no", "tier", "scenario", "capacity", "target_0", "target_1", "target_2", "edit_bound", "tour_length", "sensor_nodes", "sensor_k", "safety_vars", "rations_target"]),
    ("crates", ["case_id", "crate_id", "cat", "weight", "value", "expiry", "serial"]),
    ("locations", ["case_id", "loc_id", "type", "penalty"]),
    ("roads", ["case_id", "road_no", "u", "v", "minutes", "cable"]),
    ("waves", ["case_id", "wave_no", "road_no"]),
    ("radio_log", ["case_id", "part", "text"]),
    ("radio_patterns", ["case_id", "pattern_no", "pattern", "count"]),
    ("radio_reports", ["case_id", "report_a", "report_b"]),
    ("windows", ["case_id", "window_id", "start", "finish", "value"]),
    ("incidents", ["case_id", "alarm_id", "tick", "priority"]),
    ("sites", ["case_id", "site_no", "x", "y"]),
    ("tour", ["case_id", "step", "site_no"]),
    ("pipeline_dims", ["case_id", "idx", "dim"]),
    ("drum_prices", ["case_id", "length", "price"]),
    ("sensor_edges", ["case_id", "u", "v"]),
    ("sensor_cover", ["case_id", "node"]),
    ("safety_rules", ["case_id", "rule_no", "lit_1", "lit_2", "lit_3"]),
    ("safety_planted", ["case_id", "switch", "value"]),
    ("rations", ["case_id", "item", "value", "planted"]),
    ("claims", ["case_id", "cargo_opt", "window_count_opt", "window_value_opt", "drum_opt", "tour_opt", "cover_min"]),
]

RANGES = {
    "S": dict(crates=(8, 20), V=(6, 12), log=(200, 500), rep=(10, 30), windows=(8, 20), alarms=(10, 40), sites=(5, 9), stages=(4, 8), drum=(5, 14), nodes=(6, 14), vars=(6, 15), rations=(8, 20)),
    "M": dict(crates=(40, 200), V=(30, 150), log=(2000, 10000), rep=(30, 150), windows=(50, 300), alarms=(200, 1000), sites=(10, 16), stages=(15, 60), drum=(40, 200), nodes=(20, 40), vars=(20, 50), rations=(24, 40)),
    "L": dict(crates=(500, 1500), V=(500, 2000), log=(20000, 100000), rep=(200, 1000), windows=(1000, 3000), alarms=(2000, 5000), sites=(200, 500), stages=(100, 300), drum=(500, 2000), nodes=(100, 400), vars=(20, 50), rations=(50, 100)),
}


def cdiv(a, b):
    return -(-a // b)


def tier_of(k):
    if k < 20:
        return "S"
    if k < 40:
        return "M"
    if k < 46:
        return "L"
    return "X"


def size_of(rng, tier, key, k):
    t = "L" if tier == "X" else tier
    lo, hi = RANGES[t][key]
    if key == "crates" and k == 46:
        lo, hi = 1500, 3000
    if key == "crates" and k == 47:
        lo, hi = 12, 60
    if key == "V" and k == 48:
        lo, hi = 1500, 3000
    if key == "log" and k == 49:
        lo, hi = 50000, 100000
    return rng.randint(lo, hi)


# crates

CAT_RULES = {1: ((0, 1), 55), 2: ((2, 3), 45), 4: ((4,), 10), 6: ((2, 1), 55)}


def gen_crates(rng, sc, n):
    cats = [rng.randrange(5) for _ in range(n)]
    if sc in CAT_RULES:
        pref, pct = CAT_RULES[sc]
        need = min(n, cdiv(pct * n, 100) + rng.randint(0, max(1, n // 10)))
        for i in rng.sample(range(n), need):
            cats[i] = rng.choice(pref)
    weights = [rng.randint(1, 20) for _ in range(n)]
    values = []
    for w in weights:
        v = int(round(w * rng.uniform(2.0, 6.0) + rng.uniform(-10, 25)))
        values.append(min(100, max(10, v)))
    expiry = [rng.randint(0, 999) for _ in range(n)]
    if sc == 6:
        need = min(n, cdiv(n, 2) + rng.randint(0, max(1, n // 10)))
        for i in rng.sample(range(n), need):
            expiry[i] = rng.randint(0, 300)
    serials = rng.sample(range(10_000_000, 100_000_000), n)
    return [(i + 1, cats[i], weights[i], values[i], expiry[i], serials[i]) for i in range(n)]


def gen_crates_dock(rng, n):
    values = sorted(rng.randint(10, 100) for _ in range(n))
    weights = sorted((rng.randint(1, 20) for _ in range(n)), reverse=True)
    expiry = sorted(rng.randint(0, 999) for _ in range(n))
    serials = sorted(rng.sample(range(10_000_000, 100_000_000), n))
    cats = [rng.randrange(5) for _ in range(n)]
    return [(i + 1, cats[i], weights[i], values[i], expiry[i], serials[i]) for i in range(n)]


def pick_capacity(rng, crates):
    total = sum(c[2] for c in crates)
    lo = max(1, cdiv(total, 10))
    hi = (7 * total) // 10
    return rng.randint(lo, hi)


def greedy_value(crates, cap):
    order = sorted(crates, key=lambda c: (-Fraction(c[3], c[2]), c[0]))
    left = cap
    tot = 0
    for c in order:
        if c[2] <= left:
            left -= c[2]
            tot += c[3]
    return tot


def knap_opt(crates, cap):
    best = [0] * (cap + 1)
    for c in crates:
        w, v = c[2], c[3]
        for j in range(cap, w - 1, -1):
            if best[j - w] + v > best[j]:
                best[j] = best[j - w] + v
    return best[cap]


def gen_crates_bait(rng):
    # best-looking-crate trap: a high value-per-kg bait blocks the two crates that fill the truck exactly
    while True:
        h = rng.randint(7, 10)
        n = rng.randint(12, 16)
        p = rng.randint(2, 5)
        q = rng.randint(1, 3)
        items = [(h, 5 * h) for _ in range(p)]
        items += [(h + 1, rng.randint(5 * h + 6, 6 * h)) for _ in range(q)]
        while len(items) < n:
            w = rng.randint(h, h + 2)
            items.append((w, rng.randint(10, (49 * w) // 10)))
        total = sum(w for w, _ in items)
        if total > 20 * h:
            continue
        rng.shuffle(items)
        serials = rng.sample(range(10_000_000, 100_000_000), n)
        crates = [(i + 1, rng.randrange(5), items[i][0], items[i][1], rng.randint(0, 999), serials[i]) for i in range(n)]
        cap = 2 * h
        if 5 * greedy_value(crates, cap) <= 3 * knap_opt(crates, cap):
            return crates, cap


# network

def edge_range(sc, V):
    if sc == 1:
        return cdiv(115 * V, 100), (155 * V) // 100
    if sc == 2:
        return V, (128 * V) // 100
    if sc == 4:
        return cdiv(18 * V, 10), (26 * V) // 10
    if sc == 5:
        return V - 1, (118 * V) // 100
    return cdiv(13 * V, 10), (20 * V) // 10


def build_edges(rng, sc, V, E):
    window = max(2, {1: 10, 2: 4, 3: 12, 4: V, 5: 3, 6: 6}.get(sc, 8))
    span = {1: 12, 2: 6, 3: 15, 4: V, 5: 5, 6: 10}.get(sc, 10)
    maxdeg = 4 if sc == 5 else 10 ** 9
    deg = [0] * V
    edges = set()
    for i in range(1, V):
        lo = max(0, i - window)
        cands = [j for j in range(lo, i) if deg[j] < maxdeg]
        while not cands:
            lo = max(0, lo - window)
            cands = [j for j in range(lo, i) if deg[j] < maxdeg]
        j = rng.choice(cands)
        edges.add((j, i))
        deg[i] += 1
        deg[j] += 1
    tries = 0
    while len(edges) < E and tries < 60 * E + 1000:
        tries += 1
        a = rng.randrange(V)
        b = rng.randrange(max(0, a - span), min(V, a + span + 1))
        if a == b:
            continue
        u, v = min(a, b), max(a, b)
        if (u, v) in edges or deg[u] >= maxdeg or deg[v] >= maxdeg:
            continue
        edges.add((u, v))
        deg[u] += 1
        deg[v] += 1
    edges = sorted(edges)
    perm = list(range(1, V))
    rng.shuffle(perm)
    mp = [0] + perm
    out = []
    for u, v in edges:
        a, b = mp[u], mp[v]
        out.append((min(a, b), max(a, b)))
    rng.shuffle(out)
    return out


def build_adj(V, roads):
    adj = [[] for _ in range(V)]
    for i, (u, v, m, c) in enumerate(roads):
        adj[u].append((v, m, i))
        adj[v].append((u, m, i))
    return adj


def bfs(V, adj, src, closed):
    dist = [-1] * V
    par = [None] * V
    dist[src] = 0
    q = deque([src])
    while q:
        u = q.popleft()
        for v, m, ei in adj[u]:
            if closed and ei in closed:
                continue
            if dist[v] == -1:
                dist[v] = dist[u] + 1
                par[v] = (u, ei)
                q.append(v)
    return dist, par


def dijkstra(V, adj, pen, src):
    dist = [INF] * V
    dist[src] = 0
    pq = [(0, src)]
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


def route_trap_nodes(V, adj, pen):
    # nodes where every fewest-roads route costs more than a cheapest route
    level = [-1] * V
    level[0] = 0
    order = [0]
    q = deque([0])
    while q:
        u = q.popleft()
        for v, m, ei in adj[u]:
            if level[v] == -1:
                level[v] = level[u] + 1
                order.append(v)
                q.append(v)
    g = [INF] * V
    g[0] = 0
    for u in order:
        for v, m, ei in adj[u]:
            if level[v] == level[u] + 1 and g[u] + m + pen[v] < g[v]:
                g[v] = g[u] + m + pen[v]
    dist = dijkstra(V, adj, pen, 0)
    return [t for t in range(1, V) if g[t] > dist[t]]


def assign_types_penalties(rng, sc, V):
    frac = rng.uniform(0.28, 0.4) if sc == 3 else rng.uniform(0.15, 0.3)
    nS = max(1, round(V * frac))
    if sc == 3:
        nS = max(nS, cdiv(V, 4))
    nH = max(1, round(V * rng.uniform(0.08, 0.15)))
    while nS + nH < 3:
        nS += 1
    nodes = list(range(1, V))
    rng.shuffle(nodes)
    types = ["J"] * V
    types[0] = "D"
    for u in nodes[:nS]:
        types[u] = "S"
    for u in nodes[nS:nS + nH]:
        types[u] = "H"
    pen = [0] * V
    if sc == 1:
        need = min(V - 1, cdiv(2 * V, 5) + rng.randint(0, max(1, V // 10)))
        high = set(rng.sample(range(1, V), need))
        for u in range(1, V):
            pen[u] = rng.choice([2, 4, 6]) if u in high else rng.choice([0, 1])
    else:
        for u in range(1, V):
            pen[u] = rng.choices(PENALTIES, weights=[3, 3, 2, 2, 1])[0]
    return types, pen


def find_cut(rng, V, adj, T):
    # a target and a road whose closure cuts the target off from the HQ
    dist, par = bfs(V, adj, 0, None)
    ts = T[:]
    rng.shuffle(ts)
    for t in ts:
        path = []
        v = t
        while v != 0:
            p, ei = par[v]
            path.append(ei)
            v = p
        rng.shuffle(path)
        for ei in path[:60]:
            d2, _ = bfs(V, adj, 0, {ei})
            if d2[t] == -1:
                return t, ei
    return None


def gen_waves(rng, sc, E, force):
    w = rng.randint(3, 5) if sc in (1, 3, 5) else rng.randint(2, 5)
    cap = max(1, E // 4)
    budget = min(E, max(w, E // 2))
    sizes = []
    for i in range(w):
        left = budget - sum(sizes) - (w - 1 - i)
        mx = min(cap, left)
        if E >= 40:
            mx = min(mx, E // 8)
        sizes.append(rng.randint(1, max(1, mx)))
    waves = [[] for _ in sizes]
    if force is not None:
        waves[rng.randrange(w)].append(force)
    pool = [i for i in range(E) if i != force]
    picks = rng.sample(pool, sum(sizes) - (1 if force is not None else 0))
    pos = 0
    for i, sz in enumerate(sizes):
        need = sz - len(waves[i])
        waves[i] += picks[pos:pos + need]
        pos += need
    return [sorted(x) for x in waves]


def gen_network(rng, k, sc, tier, want_trap):
    V = size_of(rng, tier, "V", k)
    maxe = min(3 * V, V * (V - 1) // 2)
    lo, hi = edge_range(sc, V)
    lo = max(lo, V - 1)
    hi = min(hi, maxe)
    attempt = 0
    while True:
        attempt += 1
        if attempt > 3000:
            want_trap = False
        if attempt > 60000:
            raise RuntimeError("network generation failed for case %d" % k)
        E = rng.randint(lo, hi)
        edges = build_edges(rng, sc, V, E)
        if len(edges) != E:
            continue
        types, pen = assign_types_penalties(rng, sc, V)
        roads = [(u, v, rng.randint(1, 60), rng.randint(1, 100)) for u, v in edges]
        adj = build_adj(V, roads)
        dist, _ = bfs(V, adj, 0, None)
        if sc == 5 and max(dist) < V // 4:
            continue
        T = [u for u in range(V) if types[u] in "SH"]
        force = None
        if sc == 5:
            res = find_cut(rng, V, adj, T)
            if res is None:
                continue
            tcut, force = res
            others = rng.sample([t for t in T if t != tcut], 2)
            targets = [tcut] + others
            rng.shuffle(targets)
        elif want_trap:
            trap = [t for t in route_trap_nodes(V, adj, pen) if types[t] in "SH"]
            if not trap:
                continue
            tt = rng.choice(trap)
            targets = [tt] + rng.sample([t for t in T if t != tt], 2)
            rng.shuffle(targets)
        else:
            targets = rng.sample(T, 3)
        waves = gen_waves(rng, sc, E, force)
        return dict(V=V, types=types, pen=pen, roads=roads, targets=targets, waves=waves)


def gen_diamond(rng, k):
    # shortest-looks-cheapest trap: a chain of diamonds, the direct road is one hop but slow
    D = rng.randint(750, 1400)
    V = 2 * D + 1
    raw = []
    for i in range(1, D + 1):
        raw.append((i - 1, i, rng.randint(45, 60)))
        raw.append((i - 1, D + i, rng.randint(1, 5)))
        raw.append((D + i, i, rng.randint(1, 5)))
    perm = list(range(1, V))
    rng.shuffle(perm)
    mp = [0] + perm
    roads = []
    for u, v, m in raw:
        a, b = mp[u], mp[v]
        roads.append((min(a, b), max(a, b), m, rng.randint(1, 100)))
    rng.shuffle(roads)
    end = mp[D]
    types = ["J"] * V
    types[0] = "D"
    others = [u for u in range(1, V) if u != end]
    rng.shuffle(others)
    nS = V * 20 // 100
    nH = V * 10 // 100
    for u in others[:nS]:
        types[u] = "S"
    for u in others[nS:nS + nH]:
        types[u] = "H"
    types[end] = "H"
    pen = [0] + [rng.choices(PENALTIES, weights=[3, 3, 2, 2, 1])[0] for _ in range(V - 1)]
    T = [u for u in range(V) if types[u] in "SH" and u != end]
    targets = [end] + rng.sample(T, 2)
    rng.shuffle(targets)
    waves = gen_waves(rng, 0, len(roads), None)
    return dict(V=V, types=types, pen=pen, roads=roads, targets=targets, waves=waves)


# radio

def count_overlap(text, pat):
    c = 0
    i = text.find(pat)
    while i != -1:
        c += 1
        i = text.find(pat, i + 1)
    return c


def gen_log(rng, sc, N):
    weights = [1.0 / (i + 1) ** rng.uniform(0.5, 1.3) for i in range(38)]
    rng.shuffle(weights)
    if sc == 3:
        weights = [1.0 + rng.random() for _ in range(38)]
    dom = rng.randrange(38)
    if sc == 6:
        weights[dom] = 0.0
        weights[dom] = sum(weights) * 0.6
    nmsg = rng.randint(3, 25)
    msgs = [''.join(rng.choices(ALPHABET, weights, k=rng.randint(4, 14))) for _ in range(nmsg)]
    out = []
    total = 0
    while total < N:
        if rng.random() < 0.55:
            m = msgs[min(int(rng.expovariate(0.4)), nmsg - 1)]
        else:
            m = ''.join(rng.choices(ALPHABET, weights, k=rng.randint(3, 12)))
        out.append(m)
        total += len(m)
    text = list(''.join(out)[:N])
    if sc == 6:
        sym = ALPHABET[dom]
        need = cdiv(28 * N, 100)
        have = text.count(sym)
        if have < need:
            idx = [i for i, c in enumerate(text) if c != sym]
            for i in rng.sample(idx, need - have):
                text[i] = sym
    if sc == 3:
        used = set(text)
        missing = [c for c in ALPHABET if c not in used]
        rng.shuffle(missing)
        for c in missing[:max(0, 31 - len(used))]:
            text[rng.randrange(N)] = c
        # replacing may remove a symbol, so top up again from scratch
        while len(set(text)) < 31:
            used = set(text)
            c = rng.choice([x for x in ALPHABET if x not in used])
            text[rng.randrange(N)] = c
    return ''.join(text)


def gen_log_trap(rng, N):
    # near-miss trap: almost all one symbol, patterns are that symbol repeated with a wrong last symbol
    while True:
        base, absent, second = rng.sample(ALPHABET, 3)
        pool = [c for c in ALPHABET if c not in (base, absent)]
        text = [base] * N
        noise = max(20, N // 1000)
        for pos in rng.sample(range(5, N), noise):
            text[pos] = rng.choice(pool)
        for pos in rng.sample(range(5, N), 6):
            text[pos] = second
        text = ''.join(text)
        pats = [base + base, base * 3 + second] + [base * (L - 1) + absent for L in (8, 12, 20, 40)]
        if count_overlap(text, pats[1]) > 0:
            return text, pats


def gen_patterns(rng, text):
    lens = [2, 4, 8, 12, 20, 40]
    while True:
        take = set(rng.sample(range(5), rng.randint(3, 5)))
        pats = []
        for i, L in enumerate(lens):
            if i in take:
                pos = rng.randrange(len(text) - L + 1)
                pats.append(text[pos:pos + L])
            else:
                pats.append(''.join(rng.choice(ALPHABET) for _ in range(L)))
        counts = [count_overlap(text, p) for p in pats]
        if sum(1 for c in counts if c > 0) >= 2 and min(counts) == 0:
            return pats, counts


def edit_distance(a, b):
    prev = list(range(len(b) + 1))
    for i in range(1, len(a) + 1):
        cur = [i] + [0] * len(b)
        ca = a[i - 1]
        for j in range(1, len(b) + 1):
            x = prev[j - 1] + (ca != b[j - 1])
            y = prev[j] + 1
            z = cur[j - 1] + 1
            cur[j] = x if x <= y and x <= z else (y if y <= z else z)
        prev = cur
    return prev[len(b)]


def gen_reports(rng, tier, k):
    t = "L" if tier == "X" else tier
    lo, hi = RANGES[t]["rep"]
    while True:
        la = rng.randint(lo, hi)
        a = ''.join(rng.choice(ALPHABET) for _ in range(la))
        b = list(a)
        for _ in range(rng.randint(1, max(2, la // 4))):
            op = rng.randrange(3)
            pos = rng.randrange(len(b)) if b else 0
            if op == 0:
                b.insert(pos, rng.choice(ALPHABET))
            elif op == 1 and len(b) > 1:
                del b[pos]
            elif b:
                b[pos] = rng.choice(ALPHABET)
        b = ''.join(b)
        if lo <= len(b) <= hi:
            return a, b, edit_distance(a, b)


# windows and incidents

def mean_overlap(wins):
    starts = sorted(w[0] for w in wins)
    fins = sorted(w[1] for w in wins)
    tot = 0
    for s, f, _ in wins:
        tot += bisect_left(starts, f) - bisect_right(fins, s) - 1
    return Fraction(tot, len(wins))


def gen_windows(rng, sc, n):
    while True:
        hot = [rng.randint(100, 1300) for _ in range(rng.randint(2, 4))]
        wins = []
        for _ in range(n):
            if sc == 3:
                d = rng.randint(80, 480)
                s = int(rng.gauss(rng.choice(hot), 180))
                s = min(max(0, s), 1440 - d)
            else:
                d = int(rng.triangular(5, 480, 120))
                s = rng.randint(0, 1440 - d)
            wins.append([s, s + d, rng.randint(1, 100)])
        if n >= 8:
            t = rng.randint(2, 4)
            cands = [i for i in range(n) if wins[i][1] <= 1400]
            if len(cands) < t:
                continue
            a_idx = rng.sample(cands, t)
            rest = [i for i in range(n) if i not in set(a_idx)]
            b_idx = rng.sample(rest, t)
            for a, b in zip(a_idx, b_idx):
                fin = wins[a][1]
                d = rng.randint(5, min(480, 1440 - fin))
                wins[b][0] = fin
                wins[b][1] = fin + d
        if sc == 3 and mean_overlap(wins) < Fraction(23, 10):
            continue
        return [tuple(w) for w in wins]


def gen_incidents(rng, sc, tier, N):
    lo, hi = (15, 18) if sc == 4 else (12, 18)
    if tier == "S":
        t1 = rng.randint(max(2, N // 4), N)
    else:
        t1 = rng.randint(cdiv(10 * N, hi), (10 * N) // lo)
    w = [1.0] * t1
    for _ in range(rng.randint(2, 6)):
        c = rng.randrange(t1)
        amp = rng.uniform(3, 10)
        tau = rng.uniform(2, 15)
        for t in range(c, min(t1, c + int(5 * tau) + 1)):
            w[t] += amp * math.exp(-(t - c) / tau)
    ticks = [0, t1 - 1] + rng.choices(range(t1), w, k=N - 2)
    ticks.sort()
    base = [1, 2, 3, 4, 5, 6, 6, 5, 4, 3]
    burst = [1, 1, 2, 3, 4, 6, 7, 8, 8, 6]
    if sc == 2:
        base = [1, 1, 2, 2, 3, 5, 7, 8, 7, 5]
        burst = [0, 1, 1, 1, 2, 4, 7, 9, 9, 8]
    pri = []
    for t in ticks:
        ws = burst if w[t] > 4 else base
        pri.append(rng.choices(range(1, 11), ws)[0])
    if tier != "S":
        pri[rng.randrange(N)] = 1
        while len(set(pri)) < 6:
            missing = [p for p in range(1, 11) if p not in set(pri)]
            i = rng.randrange(N)
            if pri.count(pri[i]) > 1 and pri[i] != 1:
                pri[i] = rng.choice(missing)
    if sc == 2:
        while sum(pri) < 62 * N // 10:
            i = rng.randrange(N)
            if 2 <= pri[i] < 10:
                pri[i] += 1
    return [(i + 1, ticks[i], pri[i]) for i in range(N)]


# sites and tour

def manhattan(a, b):
    return abs(a[0] - b[0]) + abs(a[1] - b[1])


def tour_len(pts, order):
    n = len(order)
    return sum(manhattan(pts[order[i]], pts[order[(i + 1) % n]]) for i in range(n))


def nn_tour(pts):
    n = len(pts)
    seen = [False] * n
    seen[0] = True
    order = [0]
    while len(order) < n:
        cur = pts[order[-1]]
        best = -1
        bd = INF
        for j in range(n):
            if not seen[j]:
                d = manhattan(cur, pts[j])
                if d < bd:
                    bd = d
                    best = j
        seen[best] = True
        order.append(best)
    return order


def two_opt(pts, order, passes):
    n = len(order)
    d = manhattan
    for _ in range(passes):
        improved = False
        for i in range(1, n - 1):
            for j in range(i + 1, n):
                a, b = pts[order[i - 1]], pts[order[i]]
                c, e = pts[order[j]], pts[order[(j + 1) % n]]
                if d(a, c) + d(b, e) < d(a, b) + d(c, e):
                    order[i:j + 1] = order[i:j + 1][::-1]
                    improved = True
        if not improved:
            break
    return order


def gen_sites(rng, n):
    pts = set()
    while len(pts) < n:
        pts.add((rng.randint(0, 999), rng.randint(0, 999)))
    pts = sorted(pts)
    rng.shuffle(pts)
    return pts


def held_karp(pts):
    n = len(pts)
    dm = [[manhattan(a, b) for b in pts] for a in pts]
    full = 1 << (n - 1)
    dp = [[INF] * (n - 1) for _ in range(full)]
    for j in range(n - 1):
        dp[1 << j][j] = dm[0][j + 1]
    for mask in range(1, full):
        for j in range(n - 1):
            cur = dp[mask][j]
            if cur == INF or not (mask >> j) & 1:
                continue
            for t in range(n - 1):
                if (mask >> t) & 1:
                    continue
                nm = mask | (1 << t)
                v = cur + dm[j + 1][t + 1]
                if v < dp[nm][t]:
                    dp[nm][t] = v
    return min(dp[full - 1][j] + dm[j + 1][0] for j in range(n - 1))


# other modules

def gen_pipeline(rng, n):
    return [rng.randint(2, 100) for _ in range(n + 1)]


def gen_drum(rng, L):
    while True:
        peak = rng.randint(120, 500)
        beta = rng.uniform(0.45, 0.95)
        prices = []
        for j in range(1, L + 1):
            p = int(round(peak * (j / L) ** beta * rng.uniform(0.85, 1.1)))
            prices.append(min(500, max(1, p)))
        if any(prices[j - 1] * 1 != prices[0] * j for j in range(1, L + 1)):
            return prices


def rod_opt(prices):
    L = len(prices)
    best = [0] * (L + 1)
    for j in range(1, L + 1):
        best[j] = max(prices[i - 1] + best[j - i] for i in range(1, j + 1))
    return best[L]


def gen_sensors(rng, n):
    k = rng.randint(2, n // 2)
    cover = sorted(rng.sample(range(n), k))
    rest = [u for u in range(n) if u not in set(cover)]
    edges = set()
    for u in rest:
        c = rng.choice(cover)
        edges.add((min(u, c), max(u, c)))
    shuffled = cover[:]
    rng.shuffle(shuffled)
    for i in range(1, k):
        a, b = shuffled[i], shuffled[rng.randrange(i)]
        edges.add((min(a, b), max(a, b)))
    cand = []
    cs = set(cover)
    for u in range(n):
        for v in range(u + 1, n):
            if (u in cs or v in cs) and (u, v) not in edges:
                cand.append((u, v))
    extra = min(len(cand), rng.randint(0, n))
    for e in rng.sample(cand, extra):
        edges.add(e)
    return cover, sorted(edges)


def min_cover(n, edges):
    for size in range(n + 1):
        for comb in itertools.combinations(range(n), size):
            s = set(comb)
            if all(u in s or v in s for u, v in edges):
                return size
    return n


def gen_safety(rng, nvars, m):
    planted = [0] + [rng.randint(0, 1) for _ in range(nvars)]
    rules = []
    for r in range(1, m + 1):
        vs = rng.sample(range(1, nvars + 1), 3)
        lits = [v if rng.random() < 0.5 else -v for v in vs]
        if not any((lit > 0) == (planted[abs(lit)] == 1) for lit in lits):
            i = rng.randrange(3)
            lits[i] = -lits[i]
        rules.append((r, lits[0], lits[1], lits[2]))
    return planted, rules


def gen_rations(rng, q):
    values = [rng.randint(1, 1000) for _ in range(q)]
    size = rng.randint(max(2, q // 4), max(3, (3 * q) // 4))
    chosen = set(rng.sample(range(q), size))
    flags = [1 if i in chosen else 0 for i in range(q)]
    target = sum(values[i] for i in chosen)
    return values, flags, target


def window_opts(wins):
    ws = sorted(wins, key=lambda w: (w[1], w[0]))
    fins = [w[1] for w in ws]
    cnt = [0] * (len(ws) + 1)
    val = [0] * (len(ws) + 1)
    for i, (s, f, v) in enumerate(ws, start=1):
        p = bisect_right(fins, s)
        cnt[i] = max(cnt[i - 1], 1 + cnt[min(p, i - 1)])
        val[i] = max(val[i - 1], v + val[min(p, i - 1)])
    return cnt[-1], val[-1]


BANDS = {"low": (25, 34), "mid": (35, 48), "high": (49, 60)}


def safety_plan(s):
    bands = ["low"] * 8 + ["mid"] * 10 + ["high"] * 8
    random.Random(s * 7919 + 1).shuffle(bands)
    return bands


def make_case(s, k, plan):
    rng = random.Random(s * 1000 + k)
    tier = tier_of(k)
    sc = (k % 6) + 1
    sc_eff = sc if k < 46 else 0
    cid = "%d-%02d" % (s, k)
    rows = {}

    # crates and capacity
    n = size_of(rng, tier, "crates", k)
    if k == 46:
        crates = gen_crates_dock(rng, n)
        cap = pick_capacity(rng, crates)
    elif k == 47:
        crates, cap = gen_crates_bait(rng)
    else:
        want_fail = tier == "S" and k % 3 == 0
        for _ in range(3000):
            crates = gen_crates(rng, sc_eff, n)
            cap = pick_capacity(rng, crates)
            if not want_fail or greedy_value(crates, cap) < knap_opt(crates, cap):
                break
    rows["crates"] = [(cid,) + c for c in crates]

    # network
    want_trap = tier == "S" and (k % 6) in (0, 2, 3, 5)
    net = gen_diamond(rng, k) if k == 48 else gen_network(rng, k, sc_eff, tier, want_trap)
    V = net["V"]
    rows["locations"] = [(cid, u, net["types"][u], net["pen"][u]) for u in range(V)]
    rows["roads"] = [(cid, i, u, v, m, c) for i, (u, v, m, c) in enumerate(net["roads"])]
    rows["waves"] = [(cid, wi + 1, r) for wi, w in enumerate(net["waves"]) for r in w]
    targets = net["targets"]

    # radio
    N = size_of(rng, tier, "log", k)
    if k == 49:
        text, pats = gen_log_trap(rng, N)
        counts = [count_overlap(text, p) for p in pats]
    else:
        text = gen_log(rng, sc_eff, N)
        pats, counts = gen_patterns(rng, text)
    parts = []
    pos = 0
    while pos < len(text):
        step = rng.randint(2000, 10000) if len(text) > 10000 else rng.randint(len(text) // 3 + 1, len(text))
        parts.append(text[pos:pos + min(step, 10000)])
        pos += len(parts[-1])
    rows["radio_log"] = [(cid, i + 1, p) for i, p in enumerate(parts)]
    rows["radio_patterns"] = [(cid, i, pats[i], counts[i]) for i in range(6)]
    ra, rb, edit_bound = gen_reports(rng, tier, k)
    rows["radio_reports"] = [(cid, ra, rb)]

    # windows and incidents
    wins = gen_windows(rng, sc_eff, size_of(rng, tier, "windows", k))
    rows["windows"] = [(cid, i + 1, w[0], w[1], w[2]) for i, w in enumerate(wins)]
    inc = gen_incidents(rng, sc_eff, tier, size_of(rng, tier, "alarms", k))
    rows["incidents"] = [(cid,) + a for a in inc]

    # sites and tour
    pts = gen_sites(rng, size_of(rng, tier, "sites", k))
    if tier == "S":
        rest = list(range(1, len(pts)))
        rng.shuffle(rest)
        order = [0] + rest
    else:
        order = two_opt(pts, nn_tour(pts), 6)
    rows["sites"] = [(cid, i, p[0], p[1]) for i, p in enumerate(pts)]
    rows["tour"] = [(cid, i, o) for i, o in enumerate(order)]
    tlen = tour_len(pts, order)

    # pipeline and drum
    dims = gen_pipeline(rng, size_of(rng, tier, "stages", k))
    rows["pipeline_dims"] = [(cid, i, d) for i, d in enumerate(dims)]
    prices = gen_drum(rng, size_of(rng, tier, "drum", k))
    rows["drum_prices"] = [(cid, j + 1, p) for j, p in enumerate(prices)]

    # sensors
    sn = size_of(rng, tier, "nodes", k)
    cover, sedges = gen_sensors(rng, sn)
    rows["sensor_edges"] = [(cid, u, v) for u, v in sedges]
    rows["sensor_cover"] = [(cid, c) for c in cover]

    # safety
    nv = size_of(rng, tier, "vars", k)
    if 20 <= k < 46:
        lo10, hi10 = BANDS[plan[k - 20]]
    else:
        lo10, hi10 = 20, 60
    m = rng.randint(cdiv(lo10 * nv, 10), (hi10 * nv) // 10)
    planted, rules = gen_safety(rng, nv, m)
    rows["safety_rules"] = [(cid,) + r for r in rules]
    rows["safety_planted"] = [(cid, j, planted[j]) for j in range(1, nv + 1)]

    # rations
    rvals, rflags, rtarget = gen_rations(rng, size_of(rng, tier, "rations", k))
    rows["rations"] = [(cid, i, rvals[i], rflags[i]) for i in range(len(rvals))]

    # claims on tier S
    if tier == "S":
        wc, wv = window_opts(wins)
        rows["claims"] = [(cid, knap_opt(crates, cap), wc, wv, rod_opt(prices), held_karp(pts), min_cover(sn, sedges))]

    rows["cases"] = [(cid, s, k, tier, sc, cap, targets[0], targets[1], targets[2], edit_bound, tlen, sn, len(cover), nv, rtarget)]
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--seed", type=int, required=True, help="group number from the group-selection sheet")
    ap.add_argument("--out", default="cases.xlsx")
    ap.add_argument("--only", type=int, default=None, help="generate a single case number (debugging)")
    args = ap.parse_args()
    plan = safety_plan(args.seed)
    wb = Workbook(write_only=True)
    sheets = {}
    for name, hdr in SHEETS:
        ws = wb.create_sheet(name)
        ws.append(hdr)
        sheets[name] = ws
    ks = range(50) if args.only is None else [args.only]
    for k in ks:
        rows = make_case(args.seed, k, plan)
        for name, _ in SHEETS:
            for r in rows.get(name, []):
                sheets[name].append(list(r))
        print("case %d-%02d done" % (args.seed, k), flush=True)
    wb.save(args.out)
    print("wrote", args.out)


if __name__ == "__main__":
    main()

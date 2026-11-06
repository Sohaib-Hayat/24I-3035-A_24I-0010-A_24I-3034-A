#!/usr/bin/env python3
# builds validator/tests from a validated cases.xlsx
# usage: python validator/make_tests.py cases.xlsx
# every base case is copied into a one-case workbook (valid) and then broken in many different ways
# file names end with __Vn which is the rule family the validator must flag
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import validate as V
from openpyxl import Workbook

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tests")
SMALL = [3, 10]
TRAPS = [46, 47, 48, 49]


class State:
    def __init__(self, rows, hdr):
        self.rows = {k: [list(r) for r in v] for k, v in rows.items()}
        self.hdr = {k: list(v) for k, v in hdr.items()}
        self.order = [s[0] for s in V.SHEETS]
        self.extra = {}

    def col(self, sheet, name):
        return self.hdr[sheet].index(name)

    def get(self, sheet, name, i=0):
        return self.rows[sheet][i][self.col(sheet, name)]

    def put(self, sheet, name, value, i=0):
        self.rows[sheet][i][self.col(sheet, name)] = value

    def sort(self, sheet, key_cols):
        idx = [self.col(sheet, c) for c in key_cols]
        self.rows[sheet].sort(key=lambda r: tuple(r[i] for i in idx))


def write(path, st):
    wb = Workbook()
    wb.remove(wb.active)
    for name in st.order:
        ws = wb.create_sheet(name)
        if name in st.extra:
            for r in st.extra[name]:
                ws.append(r)
            continue
        ws.append(st.hdr[name])
        for r in st.rows[name]:
            ws.append(r)
    wb.save(path)


# mutations: (name, code, function, applies to)

def m_header(st, k):
    st.hdr["roads"][4] = "mins"


def m_swap(st, k):
    a, b = st.order.index("crates"), st.order.index("locations")
    st.order[a], st.order[b] = st.order[b], st.order[a]


def m_extra(st, k):
    st.order.append("notes")
    st.extra["notes"] = [["remark"], ["hand typed"]]


def m_caseid(st, k):
    for sheet in st.rows:
        for r in st.rows[sheet]:
            r[0] = r[0] + "x"


def m_tier(st, k):
    st.put("cases", "tier", "M" if st.get("cases", "tier") != "M" else "L")


def m_float(st, k):
    st.put("crates", "weight", 5.5)


def m_formula(st, k):
    st.put("radio_reports", "report_a", "=1+1")


def m_weight(st, k):
    st.put("crates", "weight", 25)


def m_capacity(st, k):
    st.put("cases", "capacity", sum(r[st.col("crates", "weight")] for r in st.rows["crates"]))


def m_serial(st, k):
    st.put("crates", "serial", st.get("crates", "serial", 0), 1)


def m_endpoint(st, k):
    st.put("roads", "v", 999, 0)


def m_target_hq(st, k):
    st.put("cases", "target_0", 0)


def m_disconnect(st, k):
    cid = st.rows["cases"][0][0]
    n = len(st.rows["locations"])
    st.rows["locations"].append([cid, n, "J", 0])


def m_waves(st, k):
    cid = st.rows["cases"][0][0]
    first = st.rows["waves"][0][2]
    last = max(r[1] for r in st.rows["waves"])
    st.rows["waves"].append([cid, last, first])
    st.sort("waves", ["wave_no", "road_no"])


def m_wavecount(st, k):
    st.rows["waves"] = [r for r in st.rows["waves"] if r[1] == 1]


def m_pattern(st, k):
    st.put("radio_patterns", "count", st.get("radio_patterns", "count", 0) + 1, 0)


def m_edit(st, k):
    a, b = st.rows["radio_reports"][0][1], st.rows["radio_reports"][0][2]
    st.put("cases", "edit_bound", V.edit_dist(a, b) - 1)


def m_claim(st, k):
    st.put("claims", "cargo_opt", st.get("claims", "cargo_opt") + 1)


def m_logpart(st, k):
    cid = st.rows["cases"][0][0]
    st.rows["radio_log"].append([cid, 5, "AB"])


def m_tour(st, k):
    st.put("tour", "site_no", st.get("tour", "site_no", 1), 2)


def m_tourlen(st, k):
    st.put("cases", "tour_length", st.get("cases", "tour_length") + 1)


def m_cover(st, k):
    st.rows["sensor_cover"].pop()


def m_planted(st, k):
    val = {r[1]: r[2] for r in st.rows["safety_planted"]}
    rule = st.rows["safety_rules"][0]
    for i in (2, 3, 4):
        j = abs(rule[i])
        rule[i] = -j if val[j] == 1 else j


def m_ration(st, k):
    st.put("cases", "rations_target", st.get("cases", "rations_target") + 1)


def m_sig_cat(st, k):
    for r in st.rows["crates"]:
        r[st.col("crates", "cat")] = 0


def m_sig_waves(st, k):
    st.rows["waves"] = [r for r in st.rows["waves"] if r[1] <= 2]


def m_dock(st, k):
    rnd = random.Random(5)
    for name in ("value", "weight", "expiry"):
        c = st.col("crates", name)
        vals = [r[c] for r in st.rows["crates"]]
        rnd.shuffle(vals)
        for r, v in zip(st.rows["crates"], vals):
            r[c] = v


def m_bait(st, k):
    n = len(st.rows["crates"])
    for r in st.rows["crates"]:
        r[st.col("crates", "weight")] = 1
    st.put("cases", "capacity", n // 2)


def m_direct(st, k):
    for r in st.rows["roads"]:
        r[st.col("roads", "minutes")] = 1


def m_log(st, k):
    cid = st.rows["cases"][0][0]
    rnd = random.Random(9)
    n = sum(len(r[2]) for r in st.rows["radio_log"])
    text = "".join(rnd.choice(sorted(V.ALPHA)) for _ in range(n))
    st.rows["radio_log"] = [[cid, i + 1, text[i * 10000:(i + 1) * 10000]] for i in range((n + 9999) // 10000)]


COMMON = [
    ("header_renamed", "V1", m_header),
    ("sheet_order_swapped", "V1", m_swap),
    ("extra_sheet", "V1", m_extra),
    ("bad_case_id", "V1", m_caseid),
    ("wrong_tier", "V1", m_tier),
    ("float_cell", "V1", m_float),
    ("formula_cell", "V1", m_formula),
    ("weight_out_of_range", "V2", m_weight),
    ("capacity_above_70_percent", "V2", m_capacity),
    ("wave_count_one", "V2", m_wavecount),
    ("duplicate_serial", "V3", m_serial),
    ("road_endpoint_missing", "V3", m_endpoint),
    ("target_is_hq", "V3", m_target_hq),
    ("disconnected_network", "V4", m_disconnect),
    ("waves_overlap", "V4", m_waves),
    ("wrong_pattern_count", "V5", m_pattern),
    ("edit_bound_below_true", "V5", m_edit),
    ("log_part_unnumbered", "V5", m_logpart),
    ("tour_not_permutation", "V6", m_tour),
    ("tour_length_wrong", "V6", m_tourlen),
    ("cover_misses_edge", "V6", m_cover),
    ("planted_assignment_breaks_rule", "V6", m_planted),
    ("ration_sum_wrong", "V6", m_ration),
]
SMALL_ONLY = [("wrong_claim_cargo", "V5", m_claim)]
SIG = {3: ("signature_no_rare_crates", "V7", m_sig_cat), 10: ("signature_only_two_waves", "V7", m_sig_waves)}
TRAP = {46: ("trap_dock_order_shuffled", "V8", m_dock), 47: ("trap_bait_removed", "V8", m_bait), 48: ("trap_direct_roads_cheap", "V8", m_direct), 49: ("trap_log_randomised", "V8", m_log)}


def main():
    src = sys.argv[1]
    os.makedirs(OUT, exist_ok=True)
    for f in os.listdir(OUT):
        if f.endswith(".xlsx"):
            os.remove(os.path.join(OUT, f))
    names, data, merged = V.load(src)
    hdr = {n: list(data[n][0]) for n in names}
    made = 0
    for k in SMALL + TRAPS:
        rows = {}
        for n in names:
            rows[n] = [list(r) for r in data[n][1:] if r[0].endswith("-%02d" % k)]
        base = State(rows, hdr)
        write(os.path.join(OUT, "valid_case%02d.xlsx" % k), base)
        made += 1
        if k in SMALL:
            muts = COMMON + SMALL_ONLY + [SIG[k]]
        else:
            muts = [TRAP[k]]
        for name, code, fn in muts:
            st = State(rows, hdr)
            fn(st, k)
            write(os.path.join(OUT, "case%02d_%s__%s.xlsx" % (k, name, code)), st)
            made += 1
    print("wrote %d files to %s" % (made, OUT))


if __name__ == "__main__":
    main()

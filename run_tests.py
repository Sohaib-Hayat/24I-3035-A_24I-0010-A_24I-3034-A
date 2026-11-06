#!/usr/bin/env python3
# runs the validator on every file in validator/tests
# valid_*.xlsx must PASS, every other file must be rejected and must show the rule family in its name (__Vn)
import os
import re
import subprocess
import sys

here = os.path.dirname(os.path.abspath(__file__))
tests = os.path.join(here, "tests")
files = sorted(f for f in os.listdir(tests) if f.endswith(".xlsx"))
bad = 0
print("VALIDATOR SELF-TEST (single-case copies, --partial)")
for f in files:
    p = subprocess.run([sys.executable, os.path.join(here, "validate.py"), os.path.join(tests, f), "--partial"], capture_output=True, text=True)
    codes = sorted(set(re.findall(r"^\s+(V\d)\s", p.stdout, re.M)) | set(re.findall(r"GLOBAL FAIL (V\d)", p.stdout)))
    msgs = [l.strip() for l in p.stdout.splitlines() if re.match(r"^\s+V\d\s", l)]
    if f.startswith("valid_"):
        ok = p.returncode == 0
        print("%-62s %s" % (f, "PASS (accepted)" if ok else "UNEXPECTED REJECTION"))
    else:
        want = re.search(r"__(V\d)\.xlsx$", f).group(1)
        ok = p.returncode != 0 and want in codes
        first = [m for m in msgs if m.startswith(want)]
        print("%-62s %s %s" % (f, "REJECTED" if ok else "NOT REJECTED AS EXPECTED", ",".join(codes)))
        if first:
            print("    first message: %s" % first[0])
    if not ok:
        bad += 1
        print(p.stdout[-1500:])
print("SELF-TEST RESULT: %s (%d files, %d unexpected)" % ("PASS" if bad == 0 else "FAIL", len(files), bad))
sys.exit(1 if bad else 0)

#!/usr/bin/env python3
"""
make_symbiota_upload.py -- put each specimen's jar number into the Symbiota
portal's Dynamic Properties field (Darwin Core dynamicProperties), so it is in
every export.

Reads the same label spreadsheets as make_jar_labels.py (CSV or Excel: Jar #,
Catalog_Verified / Verified, Catalog #s, ...) and a CURRENT download of the
portal (Darwin Core export zip), and writes:

  <output>.csv            catalogNumber,dynamicProperties -- for Symbiota's
                          Skeletal File Upload. Only specimens whose Dynamic
                          Properties is EMPTY in the download, so nothing that
                          is already in the portal is replaced.
  <output>_merge.csv      specimens whose Dynamic Properties already holds other
                          text: the current text and the merged value (current
                          text + jar). A skeletal upload can't change a field
                          that isn't empty, so these need the field cleared
                          first (see README).
  <output>_review.csv     everything left out, with the reason (CHECK), plus
                          specimens already carrying the same jar (info).

Left out (CHECK), so a wrong jar never reaches the portal:
  * jar number missing or uncertain ("H1339?", "H10??", "illegible");
  * jar not verified (Catalog_Verified empty) -- unless --include-unverified;
  * the same catalog number listed in two different jars;
  * a number not in the portal, or with more than one portal record;
  * a number that is a different taxon from the rest of its jar (misread tag);
  * Dynamic Properties already names a DIFFERENT jar.

Download the portal data just before running this, so the check sees what is
in the portal now.

Examples
--------
  python make_symbiota_upload.py batches/Trial_1_52_24.csv batches/Trial_2_52_24.csv \\
      --portal SymbOutput_2026-11-20_DwC-A.zip -o batches/jar_upload.csv

Requirements: the same as make_jar_labels.py (this script imports it).
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import make_jar_labels as M  # noqa: E402

FIELD = "dynamicProperties"         # Darwin Core dynamicProperties (Symbiota "Dynamic Properties")
DEFAULT_FORMAT = "jarNumber: {jar}" # value written; {jar} = the jar number
SEPARATOR = "; "                    # between existing text and the jar when merging
# finds a jar already recorded in Dynamic Properties: "jarNumber: H1334",
# "Jar: H1334", {"jarNumber":"H1334"} ...
EXISTING_JAR = re.compile(r"\bjar(?:\s*(?:number|no\.?|#))?\b\W{0,4}([A-Za-z]*\d+[A-Za-z]?(?:-\d+)?)",
                          re.I)
JAR_OK = re.compile(r"[A-Za-z]*\d+[A-Za-z]?(-\d+)?")


def merge_value(current: str, value: str, jar: str) -> str:
    """Existing text + the jar. JSON stays valid JSON ({"sex":"F"} ->
    {"sex": "F", "jarNumber": "H1334"}); anything else gets '; jarNumber: H1334'."""
    if current.lstrip().startswith("{"):
        try:
            d = json.loads(current)
            if isinstance(d, dict):
                key = value.split(":", 1)[0].strip() if ":" in value else "jarNumber"
                d[key] = jar
                return json.dumps(d, ensure_ascii=False)
        except ValueError:
            pass
    return current.rstrip(" ;") + SEPARATOR + value


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("sheets", type=Path, nargs="+",
                   help="label spreadsheets (.csv or .xlsx), as used for make_jar_labels.py")
    p.add_argument("-o", "--output", type=Path, default=Path("jar_upload.csv"))
    p.add_argument("--portal", type=Path, required=True,
                   help="a CURRENT Symbiota export (Darwin Core zip): every number is checked "
                        "against it, and existing Dynamic Properties are kept")
    p.add_argument("--jar", nargs="+", help="only these Jar #s")
    p.add_argument("--sheet", nargs="+", help="only these sheet names (Excel workbooks)")
    p.add_argument("--format", default=DEFAULT_FORMAT,
                   help=f"text written to Dynamic Properties; {{jar}} is the jar number "
                        f"(default '{DEFAULT_FORMAT}')")
    p.add_argument("--include-unverified", action="store_true",
                   help="also include jars with no Catalog_Verified (uses Catalog #s; flagged)")
    a = p.parse_args(argv)

    portal = M.load_portal(a.portal)
    print(f"Portal: {sum(len(v) for v in portal.values()):,} records from {a.portal}")
    jar_filter = {j.upper() for j in a.jar} if a.jar else None

    review: list[list] = []               # level, sheet, row, jar, catalog, issue
    candidates: list[tuple] = []          # (sheet, row, jar, catalog text)
    for path in a.sheets:
        jars, skipped = M.read_jars(path, a.sheet, jar_filter, 300, True, None, None)
        for s in skipped:
            review.append(["CHECK", s["sheet"], s["row"], s["jar"], "", s["issue"]])
        for jar in jars:
            where = [jar.sheet, jar.row, jar.jar_id]
            if not jar.jar_id or not JAR_OK.fullmatch(jar.jar_id):
                review.append(["CHECK", *where, "",
                               f"jar number missing or uncertain ('{jar.jar_id}'); "
                               "nothing uploaded for this jar"])
                continue
            unverified = any(i.startswith("not verified yet") for i in jar.problems)
            if unverified and not a.include_unverified:
                review.append(["CHECK", *where, "",
                               "not verified yet (Catalog_Verified empty); nothing uploaded "
                               "(--include-unverified to use Catalog #s)"])
                continue
            if unverified:
                review.append(["CHECK", *where, "", "not verified; uploaded from Catalog #s"])
            for c in jar.catalogs:
                if not M.NUM.match(c):
                    review.append(["CHECK", *where, c, "not a plain catalog number; left out"])
                    continue
                candidates.append((jar.sheet, jar.row, jar.jar_id, c))

    # the same number in two jars: leave all copies out
    jars_of = defaultdict(set)
    for _, _, j, c in candidates:
        jars_of[int(c)].add(j)
    clash = {n for n, js in jars_of.items() if len(js) > 1}

    # with the portal: different taxon from the rest of the jar
    names = defaultdict(Counter)
    for _, _, j, c in candidates:
        recs = portal.get(int(c), [])
        if len(recs) == 1 and M.portal_name(recs[0]):
            names[j][M.portal_name(recs[0])] += 1
    jar_majority = {j: cn.most_common(1)[0][0] for j, cn in names.items() if cn}

    rows, merges, seen = [], [], set()
    for sheet, row, j, c in candidates:
        n, where = int(c), [sheet, row, j, c]
        if n in clash:
            review.append(["CHECK", *where, "catalog number listed in more than one jar ("
                           + ", ".join(sorted(jars_of[n])) + "); left out"])
            continue
        recs = portal.get(n, [])
        if not recs:
            review.append(["CHECK", *where, "not in the portal; left out"])
            continue
        if len(recs) > 1:
            review.append(["CHECK", *where, f"{len(recs)} portal records with this number; "
                           "left out (an upload would update all of them)"])
            continue
        rec = recs[0]
        name = M.portal_name(rec)
        if j in jar_majority and name and name != jar_majority[j]:
            review.append(["CHECK", *where, f"portal says {name}, the rest of the jar is "
                           f"{jar_majority[j]}; left out (misread tag?)"])
            continue
        cat = M.clean(rec.get("catalogNumber", "")) or c
        if (cat, j) in seen:
            continue
        seen.add((cat, j))
        value = a.format.format(jar=j)
        current = M.clean(rec.get(FIELD, ""))
        if not current:
            rows.append([cat, value])
            continue
        m = EXISTING_JAR.search(current)
        if m and m.group(1).upper() == j.upper():
            review.append(["info", *where, f"Dynamic Properties already has this jar "
                           f"('{current}'); nothing to do"])
        elif m:
            review.append(["CHECK", *where, f"Dynamic Properties already names jar "
                           f"{m.group(1)} ('{current}'), not {j}; left out"])
        else:
            merges.append([cat, current, merge_value(current, value, j)])
            review.append(["info", *where, "Dynamic Properties has other text; merged value "
                           "in the _merge file (needs the field cleared before upload)"])

    a.output.parent.mkdir(parents=True, exist_ok=True)
    with a.output.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["catalogNumber", FIELD])
        w.writerows(rows)
    merge_path = a.output.with_name(a.output.stem + "_merge.csv")
    if merges:
        with merge_path.open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["catalogNumber", f"current_{FIELD}", FIELD])
            w.writerows(merges)
    rev = a.output.with_name(a.output.stem + "_review.csv")
    review.sort(key=lambda r: r[0] != "CHECK")
    with rev.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["level", "sheet", "row", "jar", "catalog", "issue"])
        w.writerows(review)

    n_jars = len({v for _, v in rows})
    print(f"{len(rows)} specimen(s) from {n_jars} jar(s) with empty Dynamic Properties "
          f"-> {a.output}  (Skeletal File Upload)")
    if merges:
        print(f"{len(merges)} specimen(s) already have other Dynamic Properties text -> "
              f"{merge_path}  (clear the field first; see README)")
    checks = [r for r in review if r[0] == "CHECK"]
    if checks:
        held = sorted({str(r[3]) for r in checks if r[3]})
        print(f"*** {len(checks)} item(s) left out (CHECK) -> {rev}"
              + (f": {', '.join(held[:20])}" + (" ..." if len(held) > 20 else "") if held else ""))
    else:
        print(f"Nothing left out -> {rev}")

if __name__ == "__main__":
    main()

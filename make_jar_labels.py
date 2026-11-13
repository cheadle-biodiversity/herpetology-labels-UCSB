#!/usr/bin/env python3
"""
make_jar_labels.py -- printer-ready jar labels for the CCBER herpetology collection.

Reads the jar-inventory workbook (one row per jar) and writes one PDF with one
label per page, laid out like example_with_collector_edit.pdf:

                 Jar#: H1176
    Batrachoseps nigriventris   (Plethodontidae)
    U.S.A., California, San Luis Obispo
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    Catalog #        Locality                     Collector
    ────────────────────────────────────────────────────────────
    40748            Cuesta Ridge                 K. Seltmann
    ...
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
          UC Santa Barbara Vertebrate Zoology – HERPETOLOGY
      Cheadle Center for Biodiversity and Ecological Restoration

Printing geometry
-----------------
The label prints left-to-right along the roll.  Every page is
  height = the paper width (default 10 cm, --paper-width), and
  length = whatever the content needs (or a fixed --length for a given jar).
If a jar has more specimens than fit down one column, the table continues in
a second (third, ...) column group to the right, so the label gets longer,
never taller.  Use --rotate if your printer driver is set to portrait.

Sizing labels to a jar (batches)
--------------------------------
Labels are printed in batches of one jar size.  Give each jar's outside
circumference and height in the sheet ("Circumference (cm)", "Height (cm)"),
or one size for the whole run with --jar-size 13x8.5.  The label for that jar
comes from jar_sizes.csv (Paper_use_Height = share of the 10 cm roll,
Paper_use_length_cm = label length); --list-jar-sizes shows it.  One PDF is
written per label size (jar_labels_5x12cm.pdf, ...).  "Label height (cm)" and
"Label length (cm)" columns, or --label-size 5x12, give a label size directly.

Every label uses the H1033/H742 format (Jar# centred over the name, one-line
footer).  Type is the largest that fits, from H1033 size (rows 10.5 pt) up to
H742 size (130%).  A label whose catalog numbers don't fit prints the count and
range plus a QR code with the full list; no other label has a QR code.  Only
labels too small even for that go below H1033 size (to 9 pt), flagged.

Saving roll: long strips (--strip)
----------------------------------
Instead of one page per label, --strip writes long pages (page width = roll
width, up to --strip-length, default 22 in) with the labels end to end down
the roll.  Two labels whose heights add up to no more than the roll width
print side by side (e.g. two 5 cm labels on the 10 cm roll).  The pairs are chosen to use the least roll.  Dashed lines mark the cuts;
the review and specimen CSVs give each jar's strip and position.

Locality and collector from the portal
--------------------------------------
  --portal SymbOutput_..._DwC-A.zip   Symbiota export; each specimen's row gets
                                      its own Locality and Collector
  Locality: jar all from one county -> the main feature of the portal locality
  ("Chiricahua Mtns."), plus the next part ("Chiricahua Mtns., Cave Creek
  Cyn.") when the label fits it at full type size.  Jar from several counties
  -> each row shows its county.  Collector: "S. Sweet", "A & B", "A et al.".
  <output>_specimens.csv lists every shortened locality/collector next to the
  full portal text for checking.  Without --portal the sheet's
  County/Locality is repeated on every row, as before.

Printing only conflict-free jars
--------------------------------
  --reconciliation RECON.xlsx   workbook made by herp_reconcile.py; only jars
                                whose status is "All matched" are printed
  --exact-names                 also skip jars whose name differs from the
                                portal (synonyms, older names, spelling)

Columns used (header names matched case-insensitively)
-------------------------------------------------------
  Jar #, Catalog #s (falls back to Side/Top Catalog), State etc.,
  County/Locality, Scientific Name, Preservative            -- current sheet
  Family, Collector                                          -- optional; add
      them to the sheet and they are used.  Without a Family column the
      family is looked up from the genus via GBIF and cached in
      family_cache.json next to this script.  Without a Collector column the
      Collector column is left off the label.

Font
----
The example uses Avenir.  ReportLab can only embed TrueType (.ttf/.ttc)
fonts, so the script looks for Avenir in the usual macOS locations; otherwise
point --font-regular / --font-italic at TrueType files, or it falls back to
Helvetica (built in, prints well on thermal heads).

Outputs
-------
  <output>.pdf          the labels
  <output>_review.csv   jars whose data needed a judgement call

Requirements:  pip install -r requirements.txt  (openpyxl, reportlab; networkx for --strip)

Examples
--------
  python make_jar_labels.py Herp_Label_Data.xlsx
  python make_jar_labels.py Herp_Label_Data.xlsx --jar H1874 H1875 -o test.pdf
  python make_jar_labels.py Batch_13x8.5.xlsx --strip
  python make_jar_labels.py Herp_Label_Data.xlsx --jar H1874 H1875 --jar-size 13x8.5
  python make_jar_labels.py Herp_Label_Data.xlsx --reconciliation Reconciliation.xlsx --exact-names
  python make_jar_labels.py Herp_Label_Data.xlsx --portal SymbOutput_DwC-A.zip --jar H539
  python make_jar_labels.py Herp_Label_Data.xlsx --portal SymbOutput_DwC-A.zip --strip
"""
from __future__ import annotations

import argparse
import csv
from datetime import date as dt_date
import difflib
import json
import math
import re
import sys
import time
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from dataclasses import dataclass, field, replace
from pathlib import Path

import openpyxl
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont, TTFontFile
from reportlab.pdfgen import canvas

DEFAULT_FOOTER = [
    "UC Santa Barbara Vertebrate Zoology – HERPETOLOGY",
    "Cheadle Center for Biodiversity and Ecological Restoration",
]

# Column headers as they appear in the workbook (lower-case, trimmed).
COL = {
    "jar": ["jar #", "jar#"],
    "catalog": ["catalog #s", "catalog #", "catalog"],
    "verified": ["catalog_verified", "catalog verified", "verified catalog #s", "verified"],
    "top": ["top catalog"],
    "side": ["side catalog"],
    "state": ["state etc.", "state"],
    "locality": ["county/locality", "locality"],
    "sciname": ["scientific name"],
    "family": ["family"],
    "collector": ["collector", "collector(s)", "collectors"],
    "preservative": ["preservative"],
    "shelf": ["shelf #letter"],
    "size": ["jar size"],
}

# ==========================================================================
# Fonts
# ==========================================================================
FONT_REG = "Helvetica"
FONT_ITAL = "Helvetica-Oblique"

AVENIR_CANDIDATES = [
    "/System/Library/Fonts/Avenir.ttc",
    "/System/Library/Fonts/Supplemental/Avenir.ttc",
    "/Library/Fonts/Avenir.ttc",
    str(Path.home() / "Library/Fonts/Avenir.ttc"),
]
AVENIR_FACES = {"regular": [b"Avenir-Book", b"AvenirLTPro-Book", b"Avenir-Roman"],
                "italic": [b"Avenir-BookOblique", b"AvenirLTPro-BookOblique",
                           b"Avenir-Oblique"]}


def _register(alias: str, path: str, index: int = 0) -> str:
    pdfmetrics.registerFont(TTFont(alias, path, subfontIndex=index))
    return alias


def setup_fonts(regular: str | None, italic: str | None) -> str:
    """Register fonts; returns a description of what is being used."""
    global FONT_REG, FONT_ITAL
    if regular:
        FONT_REG = _register("LabelRegular", regular)
        FONT_ITAL = _register("LabelItalic", italic) if italic else FONT_REG
        return f"{Path(regular).name} / {Path(italic).name if italic else '(no italic)'}"

    for path in AVENIR_CANDIDATES:
        if not Path(path).exists():
            continue
        faces = {}
        for i in range(32):
            try:
                faces[TTFontFile(path, subfontIndex=i).name] = i
            except Exception:
                break
        reg = next((faces[n] for n in AVENIR_FACES["regular"] if n in faces), None)
        ita = next((faces[n] for n in AVENIR_FACES["italic"] if n in faces), None)
        if reg is not None and ita is not None:
            FONT_REG = _register("LabelRegular", path, reg)
            FONT_ITAL = _register("LabelItalic", path, ita)
            return f"Avenir ({path})"
    return "Helvetica (Avenir not found; see --font-regular)"


# ==========================================================================
# Text clean-up
# ==========================================================================
US_STATES = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine",
    "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi",
    "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
    "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio",
    "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina",
    "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia",
    "Washington", "West Virginia", "Wisconsin", "Wyoming",
]
US_ABBR = {"CA": "California", "AZ": "Arizona", "NM": "New Mexico",
           "NV": "Nevada", "TX": "Texas", "OR": "Oregon", "FL": "Florida"}
AUS_STATES = ["Western Australia", "Northern Territory", "Queensland",
              "New South Wales", "Victoria", "South Australia", "Tasmania"]
COUNTRIES = ["Mexico", "Costa Rica", "France", "Australia", "Canada",
             "Panama", "Ecuador", "Brazil", "Madagascar", "South Africa"]
NULL_GEO = {"", "?", "n/a", "na", "no data", "none", "unlisted", "unknown",
            "out of state", "(out of state)"}
COUNTRY_US = "U.S.A."


def clean(v) -> str:
    """Cell value -> trimmed string ('' for empty). 41261.0 -> '41261'."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s.strip('"').strip()


def _fuzzy(name: str, choices: list[str], cutoff=0.85) -> str | None:
    hit = difflib.get_close_matches(name.title(), choices, n=1, cutoff=cutoff)
    return hit[0] if hit else None


def geography_header(state_raw: str, county_raw: str) -> str:
    """'CALIFORNIA' + 'San Bernardino Co.' -> 'U.S.A., California, San Bernardino'."""
    parts: list[str] = []
    st = clean(state_raw)
    if st.lower() in NULL_GEO or st.lower().startswith("mixed"):
        st = ""
    if st:
        if st.upper() in US_ABBR:
            parts = [COUNTRY_US, US_ABBR[st.upper()]]
        elif (m := _fuzzy(st, US_STATES)):
            parts = [COUNTRY_US, m]
        elif "australia" in st.lower() or _fuzzy(st, AUS_STATES):
            region = re.sub(r",?\s*australia\s*$", "", st, flags=re.I).strip(", ")
            region = _fuzzy(region, AUS_STATES) or (region.title() if region else "")
            parts = ["Australia"] + ([region] if region and region != "Australia" else [])
        elif (m := _fuzzy(st, COUNTRIES)):
            parts = [m]
        else:
            parts = [st.title() if st.isupper() else st]

    cty = clean(county_raw)
    m = re.fullmatch(r"([A-Za-z .'-]+?)\s+(Co\.?|County)", cty, flags=re.I)
    if m and parts and parts[0] == COUNTRY_US:
        name = m.group(1).strip()
        parts.append(name.title() if name.isupper() else name)
    return ", ".join(parts)


def locality_text(county_raw: str) -> str:
    s = clean(county_raw)
    return "" if s.lower() in NULL_GEO else s


def normalize_sciname(s: str) -> str:
    """'Bufo Microscaphus ' -> 'Bufo microscaphus' (genus capitalised only)."""
    s = clean(s)
    if not s:
        return ""
    words = s.split(" ")
    out = [words[0][:1].upper() + words[0][1:].lower()]
    out += [w if w.startswith("(") else w.lower() for w in words[1:]]
    return " ".join(out)


ROMAN_TOKENS = {"sp.", "spp.", "sp", "spp", "cf.", "cf", "aff.", "aff", "x", "?",
                "var.", "ssp.", "subsp."}


# ==========================================================================
# Catalog-number parsing
# ==========================================================================
NUM = re.compile(r"^\d+$")
RANGE = re.compile(r"^(\d+)\s*[-–—]\s*(\d+)$")


def parse_catalog(raw: str, max_expand: int, expand=True) -> tuple[list[str], list[str]]:
    """Return (label rows, problems).  Numbers are sorted ascending; anything
    unparseable is kept verbatim at the end and reported."""
    nums: set[int] = set()
    verbatim: list[str] = []
    problems: list[str] = []
    ranges_kept: list[tuple[int, str]] = []

    tokens = []
    for t in re.split(r"[,;]", raw):
        t = t.strip()
        if not t:
            continue
        pieces = t.split()  # "167  16575-16578" -> two tokens
        if len(pieces) > 1 and all(NUM.match(p) or RANGE.match(p) for p in pieces):
            tokens.extend(pieces)
        else:
            tokens.append(t)

    for t in tokens:
        if NUM.match(t):
            nums.add(int(t))
            continue
        m = RANGE.match(t)
        if m:
            a_s, b_s = m.groups()
            a = int(a_s)
            # shorthand: 41961-67 -> 41961-41967
            b = int(a_s[: len(a_s) - len(b_s)] + b_s) if len(b_s) < len(a_s) else int(b_s)
            if b < a:
                verbatim.append(t)
                problems.append(f"reversed range '{t}' printed as-is")
            elif b - a + 1 > max_expand:
                verbatim.append(t)
                problems.append(f"range '{t}' spans {b - a + 1} numbers (> {max_expand}); "
                                "printed as a range")
            elif expand:
                nums.update(range(a, b + 1))
            else:
                ranges_kept.append((a, f"{a}–{b}"))
            continue
        verbatim.append(t)
        problems.append(f"unparsed catalog entry '{t}' printed as-is")

    rows = [str(n) for n in sorted(nums)]
    if ranges_kept:
        rows = [s for _, s in sorted([(n, str(n)) for n in nums] + ranges_kept)]
    return rows + verbatim, problems


RANGE_SEP = " - "


def collapse_runs(rows: list[str], min_run: int = 2, sep: str = RANGE_SEP) -> list[str]:
    """Fold consecutive catalog numbers into ranges.
    ['40748', ..., '40782', '40800'] -> ['40748 - 40782', '40800'].
    Runs shorter than min_run stay as single rows; non-numeric rows pass through."""
    out: list[str] = []
    run: list[int] = []

    def flush():
        if len(run) >= max(min_run, 2):
            out.append(f"{run[0]}{sep}{run[-1]}")
        else:
            out.extend(str(n) for n in run)
        run.clear()

    for r in rows:
        if NUM.match(r):
            n = int(r)
            if run and n != run[-1] + 1:
                flush()
            run.append(n)
        else:
            flush()
            out.append(r)
    flush()
    return out


def looks_like_catalog(s: str) -> bool:
    return bool(s) and bool(re.search(r"\d{3,}", s)) and not re.search(
        r"no (catalog|category|numbers)", s, re.I)


# ==========================================================================
# Family lookup (GBIF backbone, cached)
# ==========================================================================
# Genera GBIF can't resolve on name alone (homonyms in other kingdoms/phyla)
# or misspellings it doesn't catch.  Edit freely; these win over GBIF.
FAMILY_OVERRIDES = {
    "Heterodon": "Colubridae",
    "Micrurus": "Elapidae",
    "Echiopsis": "Elapidae",
    "Eridiphas": "Colubridae",
    "Crolaus": "Viperidae",       # sheet typo for Crotalus
    "Carallus": "Boidae",         # sheet typo for Corallus
    "Ambyiroma": "Ambystomatidae",  # sheet typo for Ambystoma
}
NOT_A_GENUS = {"None", "Unlisted", "Unknown", "Data", "No"}


class FamilyLookup:
    URL = "https://api.gbif.org/v1/species/match?"

    def __init__(self, cache_path: Path, online: bool):
        self.path = cache_path
        self.online = online
        self.cache: dict = {}
        self.dirty = False
        self.failures = 0
        if cache_path.exists():
            try:
                self.cache = json.loads(cache_path.read_text())
            except Exception:
                self.cache = {}

    @staticmethod
    def genus_of(sciname: str) -> str:
        g = re.sub(r"[^A-Za-z]", "", sciname.split(" ")[0]) if sciname else ""
        return g.capitalize() if len(g) > 2 else ""

    def get(self, sciname: str) -> tuple[str, str]:
        """Return (family, note)."""
        genus = self.genus_of(sciname)
        if not genus:
            return "", "no genus to look up family"
        if not genus or genus in NOT_A_GENUS:
            return "", "no genus to look up family"
        if genus.endswith("idae"):          # sheet gave a family, not a genus
            return genus, ""
        if genus in FAMILY_OVERRIDES:
            return FAMILY_OVERRIDES[genus], ""
        if genus not in self.cache and self.online:
            hit = self._query(genus)
            if hit is not None:          # don't cache network failures
                self.cache[genus] = hit
                self.dirty = True
                self.save()              # keep progress if interrupted
        hit = self.cache.get(genus)
        if not hit:
            return "", f"family for '{genus}' not cached (run once with internet)"
        if not hit.get("family"):
            return "", f"GBIF found no family for genus '{genus}'"
        note = ""
        if hit.get("match") == "FUZZY":
            note = (f"genus '{genus}' fuzzy-matched to GBIF '{hit.get('gbif_genus')}' "
                    f"-> {hit['family']}; check spelling")
        return hit["family"], note

    def _query(self, genus: str) -> dict | None:
        q = urllib.parse.urlencode({"name": genus, "rank": "GENUS", "kingdom": "Animalia",
                                   "phylum": "Chordata"})
        req = urllib.request.Request(self.URL + q, headers={"User-Agent": "ccber-jar-labels/1.0"})
        d = None
        for attempt in range(4):
            try:
                with urllib.request.urlopen(req, timeout=15) as r:
                    d = json.load(r)
                break
            except Exception as e:  # offline, proxy, flaky connection
                err = e
                time.sleep(1.5 * (attempt + 1))
        if d is None:
            self.failures += 1
            if self.failures >= 5:
                self.online = False
                print(f"  GBIF unreachable ({err}); using cache only", file=sys.stderr)
            return None
        self.failures = 0
        ok = (d.get("kingdom") == "Animalia" and d.get("family")
              and d.get("matchType") in ("EXACT", "FUZZY") and d.get("confidence", 0) >= 80)
        return {"family": d.get("family", "") if ok else "",
                "match": d.get("matchType", ""), "gbif_genus": d.get("genus", "")}

    def save(self):
        if self.dirty:
            self.path.write_text(json.dumps(dict(sorted(self.cache.items())), indent=1))
            self.dirty = False


# ==========================================================================
# Workbook reading
# ==========================================================================
@dataclass
class Jar:
    sheet: str
    row: int
    jar_id: str
    sciname: str
    family: str
    header: str
    locality: str
    collector: str
    preservative: str
    catalogs: list[str]
    has_collector_col: bool
    problems: list[str] = field(default_factory=list)
    n_specimens: int = 0
    size_raw: str = ""
    fit_note: str = ""
    specimens: list = field(default_factory=list)   # Specimen, when --portal is given
    loc_mode: str = ""                              # "feature" or "county"
    position: str = ""                              # where it is in the PDF
    size_text: str = ""                             # jar size, for --show-size
    strip: int = 0
    jar_circ: str = ""                              # "Circumference (cm)" cell
    jar_height: str = ""                            # "Height (cm)" cell
    label_h: str = ""                               # "Label height (cm)" cell
    label_len: str = ""                             # "Label length (cm)" cell
    batch: str = ""                                 # which batch PDF the label is in


def find_header(ws, scan=10):
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=scan, values_only=True), 1):
        vals = [clean(v).lower() for v in row]
        idx = {}
        for key, names in COL.items():
            for n in names:
                if n in vals:
                    idx[key] = vals.index(n)
                    break
        if "catalog" in idx and "sciname" in idx:
            for i, v in enumerate(vals):          # jar / label dimension columns
                v = re.sub(r"[\s_.\-]+", " ", v).strip()   # Jar_cir -> "jar cir"
                if "label" in v and ("height" in v or re.search(r"\bht\b", v)):
                    idx.setdefault("label_h", i)
                elif "label" in v and ("length" in v or "width" in v):
                    idx.setdefault("label_len", i)
                elif re.match(r"(jar )?cir", v):
                    idx.setdefault("jar_circ", i)
                elif re.match(r"(jar )?(height|ht)\b", v):
                    idx.setdefault("jar_height", i)
            return r, idx
    return None, None


def load_reconciliation(path: Path, statuses: list[str], exact_names: bool) -> dict[str, set[int]]:
    """Read the 'Jars' tab of a herp_reconcile.py workbook.
    Returns {label-sheet name: {rows whose jar passes}}."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "Jars" not in wb.sheetnames:
        sys.exit(f"{path} has no 'Jars' tab; is it a herp_reconcile.py workbook?")
    rows = wb["Jars"].iter_rows(values_only=True)
    hdr = [clean(v) for v in next(rows)]
    col = {h: i for i, h in enumerate(hdr)}
    for need in ("Jar status", "Sheet", "Row"):
        if need not in col:
            sys.exit(f"{path}: 'Jars' tab has no '{need}' column")
    want = {x.lower() for x in statuses}
    ok: dict[str, set[int]] = defaultdict(set)
    for r in rows:
        if clean(r[col["Jar status"]]).lower() not in want:
            continue
        if exact_names and "Matched, name differs" in col and (r[col["Matched, name differs"]] or 0):
            continue
        ok[clean(r[col["Sheet"]])].add(int(r[col["Row"]]))
    return ok


def _recon_rows(recon: dict[str, set[int]], sheet: str) -> set[int]:
    """Reconciliation sheet names may be shortened ('Classroom Boxes')."""
    if sheet in recon:
        return recon[sheet]
    for k, v in recon.items():
        if sheet.startswith(k) or k.startswith(sheet):
            return v
    return set()


def open_sheet_file(path: Path):
    """An .xlsx workbook, or a .csv file as a one-sheet workbook named after the file."""
    if path.suffix.lower() != ".csv":
        return openpyxl.load_workbook(path, data_only=True, read_only=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = path.stem
    for enc in ("utf-8-sig", "cp1252", "latin-1"):     # Excel saves CSV as cp1252 on Windows
        try:
            with path.open(newline="", encoding=enc) as fh:
                rows = list(csv.reader(fh))
            break
        except UnicodeDecodeError:
            continue
    for r in rows:
        ws.append(r)
    return wb


def read_jars(path: Path, sheets, jar_filter, max_expand, expand, families: FamilyLookup | None,
              recon: dict[str, set[int]] | None = None):
    wb = open_sheet_file(path)
    jars: list[Jar] = []
    skipped: list[dict] = []
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        hdr_row, idx = find_header(ws)
        if hdr_row is None:
            continue
        allowed = _recon_rows(recon, ws.title) if recon is not None else None

        def get(row, key):
            i = idx.get(key)
            return clean(row[i]) if i is not None and i < len(row) else ""

        for rnum, row in enumerate(ws.iter_rows(min_row=hdr_row + 1, values_only=True),
                                   start=hdr_row + 1):
            if not any(clean(v) for v in row) or get(row, "shelf").upper() == "KEY":
                continue
            jar_id = get(row, "jar")
            if jar_filter and jar_id.upper() not in jar_filter:
                continue
            if allowed is not None and rnum not in allowed:
                continue

            sci = normalize_sciname(get(row, "sciname"))
            cat_raw = get(row, "catalog")
            notes = []
            # Catalog_Verified = what was found in the jar on inspection; it wins.
            ver_raw = get(row, "verified") if "verified" in idx else ""
            if "verified" in idx:
                if looks_like_catalog(ver_raw):
                    listed = set(parse_catalog(cat_raw, max_expand, expand)[0]) if cat_raw else set()
                    found = set(parse_catalog(ver_raw, max_expand, expand)[0])
                    missing = sorted(listed - found, key=lambda x: (len(x), x))
                    extra = sorted(found - listed, key=lambda x: (len(x), x))
                    if missing:
                        notes.append(f"not found in jar (left off label): {', '.join(missing)}")
                    if extra:
                        notes.append(f"in jar but not in Catalog #s: {', '.join(extra)}")
                    cat_raw = ver_raw
                else:
                    notes.append("not verified yet (Catalog_Verified empty); used Catalog #s")
            if not looks_like_catalog(cat_raw):
                for alt in ("side", "top"):
                    if looks_like_catalog(get(row, alt)):
                        cat_raw = get(row, alt)
                        notes.append(f"'Catalog #s' empty; used {alt.title()} Catalog")
                        break
            catalogs, problems = parse_catalog(cat_raw, max_expand, expand) if cat_raw else ([], [])

            if not sci and not catalogs:
                skipped.append({"sheet": ws.title, "row": rnum, "jar": jar_id,
                                "issue": "skipped: no scientific name and no catalog numbers"})
                continue

            family = get(row, "family")
            if not family and families and sci:
                family, note = families.get(sci)
                if note:
                    notes.append(note)

            jar = Jar(ws.title, rnum, jar_id, sci, family,
                      geography_header(get(row, "state"), get(row, "locality")),
                      locality_text(get(row, "locality")), get(row, "collector"),
                      get(row, "preservative"), catalogs, "collector" in idx,
                      notes + problems)
            if not sci:
                jar.problems.append("no scientific name")
            if not catalogs:
                jar.problems.append("no catalog numbers")
            if not jar.header:
                jar.problems.append("no usable state/country")
            if not jar_id:
                jar.problems.append("no Jar #")
            # the sheet has one locality per jar: show it once, in the header line
            loc = re.sub(r"\s+(co\.?|county)$", "", jar.locality, flags=re.I).strip()
            if loc and loc.lower() not in jar.header.lower():
                jar.header = f"{jar.header}, {loc}" if jar.header else loc
            jar.size_raw = get(row, "size")
            jar.jar_circ, jar.jar_height = get(row, "jar_circ"), get(row, "jar_height")
            jar.label_h, jar.label_len = get(row, "label_h"), get(row, "label_len")
            jars.append(jar)
    return jars, skipped


# ==========================================================================
# Portal data (Symbiota Darwin Core Archive)
# ==========================================================================
# With --portal, each catalog number's collector and locality come from the
# Symbiota export (occurrences.csv inside the DwC-A zip).
#
#   Locality: if every specimen in the jar is from one county, a short "main
#   feature" from the portal locality is printed for each specimen
#   ("Chiricahua Mtns.").  When the label has room at full type size, the
#   next part is added too ("Chiricahua Mtns., Cave Creek Cyn.").  If the jar
#   holds specimens from more than one county, each row shows its county.
#
#   "junction" is printed "jct." and "quadrant"/"quadrat" "quad." (LOCALITY_ABBREV).
#
#   Collector: "Samuel S. Sweet" -> "S. Sweet"; two people -> "A & B";
#   three or more -> "A et al."
#
# Every shortened locality and collector is written to <output>_specimens.csv
# so they can be checked against the full portal text.
@dataclass
class Specimen:
    text: str                 # catalog number as printed
    num: int | None = None
    found: bool = False
    county: str = ""          # "Cochise Co."
    state: str = ""
    country: str = ""
    locality: str = ""        # full portal locality
    loc1: str = ""            # main feature
    loc2: str = ""            # main feature + next part
    collector_raw: str = ""
    collector: str = ""
    sciname: str = ""         # portal genus + species (no authorship)
    family: str = ""


_DIST = re.compile(r"^\s*(?:~|approx\.?|ca\.?|about|appr\.?)?\s*[\d.,/ \-–]*\d\s*(?:\+\s*)?"
                   r"(?:air\s+|road\s+|rd\.?\s+)?(?:mi|mile|miles|km|m|ft|yds?|yards?)\b\.?/?\s*",
                   re.I)
_MID_DIST = re.compile(r"\s(?:~|approx\.?|ca\.?)?\s*[\d.,/ \-–]*\d\s*(?:air\s+)?"
                       r"(?:mi|mile|miles|km|m|ft|yds?)\b", re.I)
_CONNECT = re.compile(r"\s+(?:at|on|near|along|in|from|between|by|via|off)\s+", re.I)
_TRAIL = re.compile(r"\s+(?:and|of|the|at|to|on|in|near|from|with|by|&)$", re.I)
MAX_CLAUSE = 32   # longest single locality part, characters


# Words shortened on labels (applied before MAX_CLAUSE, so the space goes to
# the rest of the locality).  Add more here.
LOCALITY_ABBREV = [
    (r"\bjunction\b", "jct."),
    (r"\bqua(?:n)?dra(?:n)?t\b", "quad."),     # quadrant, quadrat (and "quandrant")
]


def _abbrev(c: str) -> str:
    for pat, short in LOCALITY_ABBREV:
        c = re.sub(pat, lambda m: short[:1].upper() + short[1:] if m.group()[:1].isupper()
                   else short, c, flags=re.I)
    return c


def _clause(c: str) -> str:
    c = _abbrev(re.sub(r"\s+", " ", c).strip(" ;:,"))
    c = _DIST.sub("", c).strip(" ;:,")
    m = _MID_DIST.search(c)
    if m and m.start() >= 6:
        c = c[:m.start()].strip(" ;:,")
    if len(c) > MAX_CLAUSE:
        cut = next((m.start() for m in _CONNECT.finditer(c) if 12 <= m.start() <= MAX_CLAUSE), None)
        if cut:
            c = c[:cut]
        else:
            out = ""
            for w in c.split(" "):
                if len(out) + len(w) + 1 > MAX_CLAUSE:
                    break
                out = f"{out} {w}".strip()
            c = out
    c = c.strip(" ;:,")
    while _TRAIL.search(c):
        c = _TRAIL.sub("", c)
    return c


def short_locality(loc: str) -> tuple[str, str]:
    """Portal locality -> (main feature, main feature + next part)."""
    if not loc:
        return "", ""
    parts = [_clause(p) for p in re.split(r"[,;]", loc)]
    parts = [p for p in parts if p and re.search(r"[A-Za-z]{2}", p)]
    if not parts:
        return "", ""
    return parts[0], ", ".join(parts[:2])


_SUFFIX = {"jr", "sr", "ii", "iii", "iv"}


def short_collector(s: str) -> str:
    """'Samuel S. Sweet' -> 'S. Sweet'; 'Jennings, William' -> 'W. Jennings'."""
    if not s:
        return ""
    people = [p.strip(" ,") for p in re.split(r";|\s+&\s+|\s+and\s+|\|", s) if p.strip(" ,")]
    out = []
    for p in people:
        p = re.sub(r"\s+", " ", p)
        if "," in p:
            last, first = [x.strip() for x in p.split(",", 1)]
        else:
            w = p.split(" ")
            if len(w) > 1 and w[-1].rstrip(".").lower() in _SUFFIX:
                w = w[:-1]
            last, first = w[-1], " ".join(w[:-1])
        out.append((first[0].upper() + ". " if first[:1].isalpha() else "") + last)
    if len(out) == 1:
        return out[0]
    if len(out) == 2:
        return f"{out[0]} & {out[1]}"
    return f"{out[0]} et al."


def _county(s: str) -> str:
    s = clean(s)
    if not s:
        return ""
    s = re.sub(r"\s+(County|Co\.?)$", "", s, flags=re.I).strip()
    return f"{s} Co."


def load_portal(path: Path) -> dict[int, list[dict]]:
    """Read occurrences.csv from a DwC-A zip, an unzipped folder, or the csv itself.
    Returns {catalog number: [records]}."""
    import io
    import zipfile
    raw = None
    if path.is_file() and path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            name = next((n for n in z.namelist()
                         if Path(n).name.lower() in ("occurrences.csv", "occurrence.csv")), None)
            if name:
                raw = z.read(name)
    elif path.is_dir():
        for n in ("occurrences.csv", "occurrence.csv"):
            if (path / n).exists():
                raw = (path / n).read_bytes()
                break
    elif path.is_file():
        raw = path.read_bytes()
    if raw is None:
        sys.exit(f"No occurrences.csv found in {path}")
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    out: dict[int, list[dict]] = defaultdict(list)
    for rec in csv.DictReader(io.StringIO(text)):
        m = re.match(r"\s*(\d+)", rec.get("catalogNumber") or "")
        if m:
            out[int(m.group(1))].append(rec)
    return out


def portal_name(r: dict) -> str:
    """Genus and species only, from the portal record: 'Pituophis catenifer'.
    Authorship (scientificNameAuthorship, e.g. 'Klauber 1902') and subspecies
    are left off; a species hybrid keeps both parents ('Thamnophis atratus x
    hammondii')."""
    genus = clean(r.get("genus", ""))
    ep = clean(r.get("specificEpithet", ""))
    words = clean(r.get("scientificName", "")).split()
    if not genus and words:
        genus = words[0]
    if not ep and len(words) > 1 and re.fullmatch(r"[a-z][a-z\-]+", words[1]):
        ep = words[1]
    if not genus:
        return ""
    name = f"{genus[:1].upper()}{genus[1:].lower()}" + (f" {ep.lower()}" if ep else "")
    if ep and len(words) >= 4 and words[1].lower() == ep.lower() and words[2].lower() in ("x", "\u00d7"):
        name += f" x {words[3].lower()}"
    return name


def attach_portal(jar: Jar, portal: dict[int, list[dict]]):
    """Fill jar.specimens from the portal and decide county vs feature localities."""
    specs = []
    missing = []
    for t in jar.catalogs:
        s = Specimen(t, int(t) if NUM.match(t) else None)
        recs = portal.get(s.num, []) if s.num is not None else []
        if recs:
            r = recs[0]
            s.found = True
            s.county = _county(r.get("county", ""))
            s.state = clean(r.get("stateProvince", ""))
            s.country = clean(r.get("country", ""))
            s.locality = clean(r.get("locality", ""))
            s.loc1, s.loc2 = short_locality(s.locality)
            s.collector_raw = clean(r.get("recordedBy", ""))
            s.collector = short_collector(s.collector_raw)
            s.sciname = portal_name(r)
            s.family = clean(r.get("family", ""))
            if len(recs) > 1:
                jar.problems.append(f"{t}: {len(recs)} portal records; used the first")
        elif s.num is not None:
            missing.append(t)
        specs.append(s)
    if missing:
        jar.problems.append(f"not in portal (no locality/collector): {', '.join(missing[:10])}"
                            + (" ..." if len(missing) > 10 else ""))
    found = [s for s in specs if s.found]
    # Scientific name: the portal's genus + species, not the jar's old label.
    names = Counter(s.sciname for s in found if s.sciname)
    if names:
        top, n_top = names.most_common(1)[0]
        old = jar.sciname
        if len(names) > 1:
            other = [f"{s.text} {s.sciname}" for s in found if s.sciname and s.sciname != top]
            jar.problems.append(f"specimens have different names in the portal; label uses {top} "
                                f"({n_top} of {sum(names.values())}); others: " + ", ".join(other))
        if top.lower() != old.lower():
            jar.problems.append(f"name from portal: {top} (sheet: {old or 'none'})")
        jar.sciname = top
        fams = Counter(s.family for s in found if s.sciname == top and s.family)
        if fams:
            jar.family = fams.most_common(1)[0][0]
    counties = {s.county for s in found if s.county}
    jar.specimens = specs
    jar.loc_mode = "county" if len(counties) > 1 else "feature"
    # header from the portal when the jar is one state (and county)
    states = {s.state for s in found if s.state}
    countries = {s.country for s in found if s.country}
    if len(states) == 1 and len(countries) <= 1:
        country = next(iter(countries), "")
        cty = COUNTRY_US if country.lower() in ("united states", "usa", "u.s.a.", "us", "") else country
        parts = [cty, next(iter(states))]
        if len(counties) == 1:
            parts.append(next(iter(counties)).removesuffix(" Co."))
        jar.header = ", ".join(parts)


# ==========================================================================
# Layout
# ==========================================================================
# One label design in two styles:
#
#   full      Jar#: H1176                                         ┌──┐
#             Batrachoseps nigriventris   (Plethodontidae)        │QR│
#             U.S.A., California, San Luis Obispo                 └──┘
#             ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#             Catalog #     Locality               Collector
#             ────────────────────────────────────────────────────────
#             40748         Cuesta Ridge           K. Seltmann
#             ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#                UC Santa Barbara Vertebrate Zoology – HERPETOLOGY
#              Cheadle Center for Biodiversity and Ecological Restoration
#
#   compact   (smaller jars) tighter margins, no column headings, a one-line
#             footer "UCSB Vertebrate Zoology – HERPETOLOGY – CCBER"
#             (wrapped only if the label is too narrow for it), the family
#             and the species epithet move to their own lines if needed.
#
# Type never goes below H1033 size (specimen rows --min-pt, default 10.5 pt;
# the footer is the same size as the geography line).  When a jar's table
# can't fit at that size the label falls back, in order, to:
#   numbers   catalog numbers only, in columns (locality/collector left off)
#   summary   "N specimens: first – last" and "Full list: scan QR code"
#   minimal   as summary, without the geography line, family and footer
# All three are flagged in the review file.
#
# Only the summary and minimal labels get a QR code (top right): it holds the jar
# number and every catalog number as plain text, one per line.
#
# If the rows don't fit down one column, the table continues in further
# column groups to the right, so the label gets longer, never taller.
COMPACT_FOOTER = ["UCSB Vertebrate Zoology – HERPETOLOGY – CCBER"]
# Type range (6 Nov 2026): every label sits between the two trial labels that
# read best -- Jar H1033 (Jar# 11 pt, name 13 pt, geography/rows/footer 10.5 pt)
# at the small end and Jar H742 (the same at 130%) at the large end.  Each label
# gets the largest type in that range that fits it.
MIN_PT = 10.5         # smallest specimen-row type (pt) = H1033
MAX_SCALE = 1.3       # largest type = H742 (130% of H1033)
SMALLEST_PT = 9.0     # only for labels too small for H1033 type even in count-and-range form
SIZE_PT = 6.0         # type size of the --show-size note
DOT_MM = 25.4 / 203   # one dot of a 203 dpi print head
QR_DOTS = 3           # QR module = 3 dots (0.375 mm); 2-dot codes did not scan reliably
QR_LEVEL = "L"        # error correction; L keeps the code smallest
QR_COLLECTION = "UCSB Herpetology"
QR_PREFIX = "UCSB-HERP"   # put before every catalog number in the QR code
QR_RANGES_OVER = 15       # jars with more specimens than this list ranges in the QR code
QR_DATE = dt_date.today().isoformat()      # "Last update" column; --qr-date overrides
QR_GAP = 0.25 * cm    # space between the header text and the QR code


@dataclass
class Style:
    scale: float = 1.0
    footer: list[str] = field(default_factory=lambda: list(DEFAULT_FOOTER))
    compact_footer: list[str] = field(default_factory=lambda: list(COMPACT_FOOTER))
    compact: bool = False
    valign: str = "center"
    show_preservative: bool = False
    min_run: int = 2
    ranges: bool = True
    min_pt: float = MIN_PT
    show_size: bool = False
    qr: bool = True
    qr_dots: int = QR_DOTS
    thick: float = 1.5   # pt, matches the docx border weights
    thin: float = 0.75

    def s(self, pt): return pt * self.scale
    @property
    def margin_x(self): return (0.3 if self.compact else 0.5) * cm
    @property
    def margin_y(self): return (0.25 if self.compact else 0.3) * cm
    @property
    def name_size(self): return self.s(16 if not self.compact else 13)
    @property
    def jar_size(self): return self.s(16 if not self.compact else 11)
    @property
    def geo_size(self): return self.s(14 if not self.compact else 10.5)
    @property
    def head_size(self): return self.s(12)
    @property
    def row_size(self): return self.s(10.5)
    @property
    def row_pitch(self): return self.row_size * (1.4 if not self.compact else 1.3)
    @property
    def footer_size(self): return self.geo_size        # same as the geography line
    @property
    def rule(self): return self.thick if not self.compact else 1.2


def name_runs(name: str):
    """(text, font) runs: italic name, but 'sp.', 'cf.' etc. upright."""
    return [((" " if i else "") + w, FONT_REG if w.lower() in ROMAN_TOKENS else FONT_ITAL)
            for i, w in enumerate(name.split(" "))]


def runs_width(runs, size):
    return sum(stringWidth(t, f, size) for t, f in runs)


def fit_size(text, font, size, max_w, min_size):
    while size > min_size and stringWidth(text, font, size) > max_w:
        size -= 0.25
    return size


def wrap_words(text: str, font: str, size: float, max_w: float | None) -> list[str]:
    """Wrap to max_w, preferring breaks after ' – ' so a line never starts with a dash
    and a short last part ('CCBER') doesn't sit on its own line."""
    if max_w is None or stringWidth(text, font, size) <= max_w:
        return [text]
    parts = text.split(" – ")
    if len(parts) > 1:
        best = None
        for k in range(1, len(parts)):        # one break, at a dash
            a, b = " – ".join(parts[:k]) + " –", " – ".join(parts[k:])
            if stringWidth(a, font, size) <= max_w and stringWidth(b, font, size) <= max_w:
                score = abs(stringWidth(a, font, size) - stringWidth(b, font, size))
                best = min(best, (score, [a, b])) if best else (score, [a, b])
        if best:
            return best[1]
    lines, cur = [], ""
    for w in text.split(" "):
        t = f"{cur} {w}".strip()
        if cur and stringWidth(t, font, size) > max_w:
            lines.append(cur)
            cur = w
        else:
            cur = t
    return lines + ([cur] if cur else [])


def table_rows(jar: Jar, level: int, st: Style) -> list[tuple[str, str, str]]:
    """(catalog, locality, collector) rows.  Consecutive catalog numbers with the
    same locality and collector fold into a range ('40748 - 40782')."""
    if not jar.specimens:   # no portal data: one locality for the jar, already in the header
        cats = collapse_runs(jar.catalogs, st.min_run) if st.ranges else jar.catalogs
        return [(c_, "", jar.collector) for c_ in cats]

    def loc(s: Specimen):
        if jar.loc_mode == "county":
            return s.county
        return (s.loc2 or s.loc1) if level >= 2 else s.loc1

    rows: list[tuple[str, str, str]] = []
    run: list[Specimen] = []

    def flush():
        if not run:
            return
        if st.ranges and len(run) >= max(st.min_run, 2):
            rows.append((f"{run[0].num}{RANGE_SEP}{run[-1].num}", loc(run[0]), run[0].collector))
        else:
            rows.extend((s.text, loc(s), s.collector) for s in run)
        run.clear()

    for s in jar.specimens:
        if (s.num is not None and run and s.num == run[-1].num + 1
                and loc(s) == loc(run[-1]) and s.collector == run[-1].collector):
            run.append(s)
            continue
        flush()
        if s.num is None:
            rows.append((s.text, "", ""))
        else:
            run.append(s)
    flush()
    return rows


# ---- QR code ------------------------------------------------------------
_QR_CACHE: dict[str, int] = {}


def qr_text(jar: Jar) -> str:
    """Plain text, one catalog number per line:
        Jar H1033 UCSB Herpetology
        Updated 2026-10-23
        11 specimens:
        UCSB-HERP 8252
        ...
    Jars with more than QR_RANGES_OVER specimens list ranges (UCSB-HERP 10201-10212)
    to keep the code small enough to scan reliably."""
    cats = jar.catalogs
    if len(cats) > QR_RANGES_OVER:
        cats = collapse_runs(cats, 2, "-")
    n = len(jar.catalogs)
    head = f"Jar {jar.jar_id} {QR_COLLECTION}" if jar.jar_id else QR_COLLECTION
    lines = [head, f"Updated {QR_DATE}", f"{n} specimen{'s' if n != 1 else ''}:"]
    return "\n".join(lines + [f"{QR_PREFIX} {c_}" if QR_PREFIX else c_ for c_ in cats])


def qr_modules(text: str) -> int:
    if text not in _QR_CACHE:
        from reportlab.graphics.barcode.qr import QrCodeWidget
        w = QrCodeWidget(text, barLevel=QR_LEVEL, barBorder=0)
        w.draw()
        _QR_CACHE[text] = w.qr.getModuleCount()
    return _QR_CACHE[text]


def draw_qr(c: canvas.Canvas, text: str, x: float, y: float, size: float):
    """QR code with its bottom-left corner at (x, y), size x size points."""
    from reportlab.graphics.barcode.qr import QrCodeWidget
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics import renderPDF
    w = QrCodeWidget(text, barLevel=QR_LEVEL, barBorder=0)
    b = w.getBounds()
    bw, bh = b[2] - b[0], b[3] - b[1]
    d = Drawing(size, size, transform=[size / bw, 0, 0, size / bh, 0, 0])
    d.add(w)
    renderPDF.draw(d, c, x, y)


# ---- layout ---------------------------------------------------------------
@dataclass
class Column:
    title: str
    width: float
    values: list[str]
    size: float


@dataclass
class Layout:
    length: float
    height: float
    content_h: float
    style: Style
    level: int
    mode: str                      # "table", "numbers", "summary" or "minimal"
    columns: list[Column]
    group_w: float
    group_gap: float
    groups: int
    rows_per_group: int
    name_lines: list               # [[runs], ...]  (1 or 2 lines)
    family_text: str
    family_own_line: bool
    jar_text: str
    header_h: float                # height of the header block (text or QR, whichever is taller)
    footer_lines: list[str]
    qr_text: str = ""
    qr_size: float = 0.0
    show_geo: bool = True
    natural_length: float = 0.0    # length the content needs
    fits_height: bool = True
    fits: bool = True
    qr_side: bool = False          # QR in a column down the right side
    name_pt: float = 0.0
    jar_pt: float = 0.0
    geo_pt: float = 0.0
    fam_pt: float = 0.0


def plan(jar: Jar, st: Style, paper_h: float, max_len: float | None, level: int = 1,
         mode: str = "table", trim: bool = True, qr_place: str | None = None) -> Layout:
    """Lay out a label paper_h tall.  max_len: the longest the label may be (the
    jar's 85% of circumference), or None.  With trim the label is cut to its
    content; otherwise it is max_len long (or the content, if longer)."""
    mx, my = st.margin_x, st.margin_y
    avail_h = paper_h - 2 * my
    avail_w = max_len - 2 * mx if max_len else None

    # --- QR code (its size depends on the header height, below)
    qtext = qr_text(jar) if st.qr and jar.catalogs else ""
    q_min = qr_modules(qtext) * st.qr_dots * DOT_MM / 10 * cm if qtext else 0.0

    # --- header lines
    minimal = mode == "minimal"
    show_geo = bool(jar.header) and not minimal
    name = jar.sciname or "—"
    words = name.split(" ")
    name_size, jar_size, geo_size = st.name_size, st.jar_size, st.geo_size
    family_inline = f"    ({jar.family})" if jar.family and not minimal else ""
    jar_text = f"Jar#: {jar.jar_id}" if jar.jar_id else ""
    fam_sz = st.s(12 if not st.compact else 10.5)

    def header_height(n_name_lines, fam_own):
        return ((jar_size * 1.25 if jar_text else 0) + n_name_lines * name_size * 1.25
                + (fam_sz * 1.25 if fam_own else 0) + (geo_size * 1.3 if show_geo else 0))

    def text_room(hh):
        q = q_min if qtext else 0
        return (avail_w - (q + QR_GAP if q else 0)) if avail_w else None, q

    name_lines = [name_runs(name)]
    family_text, fam_own = family_inline, False
    hh = header_height(1, False)
    room, q = text_room(hh)
    if room is not None and runs_width(name_lines[0], name_size) + \
            stringWidth(family_inline, FONT_REG, name_size) > room:
        family_text, fam_own = (f"({jar.family})" if family_inline else ""), bool(family_inline)
        hh = header_height(1, fam_own)
        room, q = text_room(hh)
        if st.compact and len(words) > 1 and runs_width(name_lines[0], name_size) > room:
            name_lines = [name_runs(words[0]), name_runs(" ".join(words[1:]))]
            hh = header_height(2, fam_own)
            room, q = text_room(hh)
    # compact labels: header lines may shrink to the readable minimum to fit beside the QR
    if st.compact and room is not None:
        floor = st.min_pt
        while name_size > floor and max(runs_width(r, name_size) for r in name_lines) > room:
            name_size -= 0.25
        jar_size = fit_size(jar_text, FONT_REG, jar_size, room, floor)
        if show_geo:
            geo_size = fit_size(jar.header, FONT_REG, geo_size, room, floor)
        if fam_own:
            fam_sz = fit_size(family_text, FONT_REG, fam_sz, room, floor)
        hh = header_height(len(name_lines), fam_own)
    # A QR code taller than the header text goes in a column down the right side
    # instead of pushing the table down (or beside the header, whichever fits).
    if qr_place is None:
        side = bool(q) and q > hh + 1
    else:
        side = bool(q) and qr_place == "side"
    header_h = hh if side else max(hh, q)
    side_w = (q + QR_GAP) if side else 0.0

    fam_size = fam_sz if fam_own else name_size
    head_w = max([runs_width(r, name_size) for r in name_lines] +
                 [stringWidth(jar_text, FONT_REG, jar_size),
                  stringWidth(jar.header, FONT_REG, geo_size) if show_geo else 0])
    if fam_own:
        head_w = max(head_w, stringWidth(family_text, FONT_REG, fam_size))
    else:
        head_w = max(head_w, runs_width(name_lines[0], name_size)
                     + stringWidth(family_text, FONT_REG, name_size))
    head_w += (q + QR_GAP) if (q and not side) else 0

    # --- footer
    if minimal:
        footer_lines = []
    elif st.compact:
        footer_lines = []
        for t in st.compact_footer:
            footer_lines += wrap_words(t, FONT_REG, st.footer_size,
                                       avail_w - side_w if avail_w else None)
    else:
        footer_lines = list(st.footer)
    foot_w = max((stringWidth(t, FONT_REG, st.footer_size) for t in footer_lines), default=0)
    if not st.compact:
        foot_w += 1.0 * cm
    if st.show_preservative and jar.preservative and not st.compact:
        foot_w += 2 * (stringWidth(jar.preservative, FONT_REG, st.s(8)) + 0.5 * cm)

    # --- table
    if mode in ("summary", "minimal"):
        nums = [s for s in jar.catalogs if NUM.match(s)]
        span = f"{nums[0]} – {nums[-1]}" if len(nums) > 1 else (nums[0] if nums else "")
        n = len(jar.catalogs)
        rows = [(f"{n} specimen{'s' if n != 1 else ''}" + (f": {span}" if span else ""), "", ""),
                ("Full list: scan QR code" if qtext else "Full list: see catalog", "", "")]
    else:
        rows = table_rows(jar, level, st) or [("—", "", "")]
        if mode == "numbers":
            rows = [(r[0], "", "") for r in rows]
    show_heads = not st.compact and mode == "table"
    cols = [Column("Catalog #", 0, [r[0] for r in rows], st.row_size)]
    if mode == "table" and any(r[1] for r in rows):
        cols.append(Column("Locality", 0, [r[1] for r in rows], st.row_size))
    if mode == "table" and (any(r[2] for r in rows) or jar.has_collector_col):
        cols.append(Column("Collector", 0, [r[2] for r in rows], st.row_size))
    pad = (0.35 if st.compact else 0.7) * cm
    for i, col in enumerate(cols):
        w = max([stringWidth(col.title, FONT_REG, st.head_size) if show_heads else 0] +
                [stringWidth(v, FONT_REG, col.size) for v in col.values])
        last = i == len(cols) - 1
        col.width = w + (0.1 * cm if last else pad)
        if i == 0 and show_heads:
            col.width = max(col.width, 2.6 * cm)
    group_w = sum(col.width for col in cols)
    gap = (0.3 if mode == "numbers" else 0.5 if st.compact else 1.2) * cm

    heads_h = st.head_size * 1.6 + 3 if show_heads else 0
    size_h = 0          # the --show-size note sits in the top margin
    fixed_v = (size_h + header_h + 4 + heads_h + 4
               + (3 + len(footer_lines) * st.footer_size * 1.2 + 2 if footer_lines else 2))
    fit_rows = max(int((avail_h - fixed_v) // st.row_pitch), 1)
    n = len(rows)
    groups = math.ceil(n / fit_rows)
    per = math.ceil(n / groups)
    table_w = groups * group_w + (groups - 1) * gap

    natural = max(table_w, head_w, foot_w) + side_w + 2 * mx
    if trim or not max_len:
        length = natural
    else:
        length = max(natural, max_len)
    content_h = max(fixed_v + per * st.row_pitch, q if side else 0)
    fits_h = content_h <= avail_h + 0.5
    fits = fits_h and (max_len is None or natural <= max_len + 0.5)
    L = Layout(length, paper_h, content_h, st, level, mode, cols, group_w, gap, groups, per,
                  name_lines, family_text, fam_own, jar_text, header_h, footer_lines,
                  qtext, q, show_geo, natural, fits_h, fits, side,
                  name_size, jar_size, geo_size, fam_sz)
    if qr_place is None and side:
        # also try the QR beside the header (pushes the table down a little);
        # keep whichever fits, preferring the shorter label
        alt = plan(jar, st, paper_h, max_len, level, mode, trim, "header")
        if alt.fits and (not L.fits or alt.natural_length < L.natural_length):
            return alt
    return L


# ==========================================================================
# Fitting a label to a jar
# ==========================================================================
@dataclass
class Fit:
    layout: Layout
    fits: bool

    @property
    def kind(self):
        L = self.layout
        return ("compact" if L.style.compact else "full") + ("" if L.mode == "table" else f", {L.mode}")

    @property
    def row_pt(self): return self.layout.style.row_size


def _best(jar, st, box_h, box_len, level, compact, mode):
    """Largest type scale (MAX_SCALE down to the readable minimum) that fits."""
    lo = st.min_pt / 10.5
    scales = []
    s_ = MAX_SCALE
    while s_ > lo + 1e-9:
        scales.append(s_)
        s_ = round(s_ - 0.025, 3)
    scales.append(lo)                       # always try the readable minimum itself
    for scale in scales:
        L = plan(jar, replace(st, scale=scale, compact=compact), box_h, box_len, level, mode)
        if L.fits:
            return L
    return None


_MODE_RANK = {"table": 0, "numbers": 1, "summary": 2, "minimal": 3}


def fit_label(jar: Jar, st: Style, box_h: float, box_len: float, min_scale: float = 0.0) -> Fit:
    """Choose style, content and type size for a jar's label box (box_len is the
    longest allowed; the label is trimmed to its content)."""
    return _fit_once(jar, st, box_h, box_len)


def _fit_once(jar: Jar, st: Style, box_h: float, box_len: float) -> Fit:
    """Every label uses the H1033/H742 (compact) format.  No QR code while the
    catalog numbers fit; the code only appears on the count-and-range fallbacks,
    where it carries the full list."""
    plain = replace(st, qr=False)
    pick = _best(jar, plain, box_h, box_len, 1, True, "table")
    if pick is not None:
        # longer localities when they cost nothing in type size
        if jar.specimens and jar.loc_mode == "feature":
            L2 = plan(jar, pick.style, box_h, box_len, 2, "table")
            if L2.fits:
                pick = L2
        return Fit(pick, True)
    for mode in ("numbers", "summary", "minimal"):
        L = _best(jar, plain if mode == "numbers" else st, box_h, box_len, 1, True, mode)
        if L:
            return Fit(L, True)
    if st.min_pt > SMALLEST_PT:
        # The smallest vials can't take H1033-size type at all: go below it, down
        # to SMALLEST_PT, rather than print a label longer than the jar (flagged).
        small = replace(st, min_pt=SMALLEST_PT)
        for mode in ("summary", "minimal"):
            L = _best(jar, small, box_h, box_len, 1, True, mode)
            if L:
                return Fit(L, True)
    # Nothing fits even minimal: readable type, label made longer, flagged.
    s2 = replace(st, scale=st.min_pt / 10.5, compact=True)
    return Fit(plan(jar, s2, box_h, None, 1, "minimal"), False)


# ==========================================================================
# Jar sizes
# ==========================================================================
# A jar is described by its outside CIRCUMFERENCE and HEIGHT in cm.
#
# Size classes A-N were measured on jars in the collection (Jar_Sizes.xlsx).
# Where several jars were measured for one class, the smallest circumference
# and smallest height are used, so the label fits every jar in that class.
# Put the letter in the sheet's "Jar Size" column, or load an updated table
# with --jar-table Jar_Sizes.xlsx (columns Circumference_cm, Height_cm,
# Size_cat_cm).
@dataclass
class JarSpec:
    key: str
    description: str
    circumference_cm: float
    height_cm: float
    estimated: bool = False


def _cls(letter, circ, height, measured):
    return JarSpec(letter, f"class {letter} (measured: {measured})", circ, height)


JAR_SIZES: dict[str, JarSpec] = {j.key: j for j in [
    _cls("A", 6.5, 7.0, "6.5 x 7"),
    _cls("B", 7.2, 9.0, "7.2 x 9"),
    _cls("C", 8.5, 9.5, "8.5 x 9.5, 8.6 x 11"),
    _cls("D", 10.7, 7.0, "10.7 x 7"),
    _cls("E", 13.0, 8.5, "13 x 8.5, 13.2 x 8.5"),
    _cls("F", 15.0, 9.0, "15 x 10.05, 16 x 9, 16 x 10.3"),
    _cls("G", 17.3, 7.0, "17.3 x 7"),
    _cls("H", 17.6, 13.7, "17.6 x 13.7"),
    _cls("I", 19.5, 11.5, "19.5 x 11.5, 19.5 x 12.7"),
    _cls("J", 23.0, 9.0, "23 x 9"),
    _cls("K", 23.7, 14.5, "23.7 x 14.5, 24 x 15"),
    _cls("L", 28.5, 9.5, "28.5 x 9.5, 28.7 x 9.5, 29.2 x 9.5"),
    _cls("M", 30.0, 17.0, "30 x 17"),
    _cls("N", 52.0, 24.0, "52 x 24"),
]}

# Thermo Scientific wide-mouth tall jars, kept from the first version of this
# script (diameter x height in mm from thermofisher.com; 500 mL estimated).
for _k, _desc, _d, _h, _est in [
        ("125ml", "125 mL (4 oz) Thermo wide-mouth", 51, 102, False),
        ("250ml", "250 mL (8 oz) Thermo wide-mouth", 62, 127, False),
        ("500ml", "500 mL (16 oz) Thermo wide-mouth", 76, 150, True),
        ("950ml", "950 mL (32 oz) Thermo wide-mouth", 95, 170, False),
        ("2000ml", "2 L (64 oz) Thermo wide-mouth", 124, 214, False),
        ("4000ml", "4 L (1 gal) Thermo wide-mouth", 133, 253, False)]:
    JAR_SIZES[_k] = JarSpec(_k, _desc, round(math.pi * _d / 10, 1), _h / 10, _est)

# Other names accepted on the command line or in the "Jar Size" column.
# small/medium/large are a GUESS at what the data-entry team meant; once jars
# have size letters, point these at letters instead.
JAR_ALIASES = {
    "4oz": "125ml", "8oz": "250ml", "16oz": "500ml", "pint": "500ml",
    "32oz": "950ml", "1l": "950ml", "quart": "950ml",
    "64oz": "2000ml", "2l": "2000ml", "halfgallon": "2000ml",
    "128oz": "4000ml", "4l": "4000ml", "gallon": "4000ml", "1gallon": "4000ml",
    "xxsmall": "125ml", "xsmall": "125ml", "small": "250ml", "medium": "500ml",
    "large": "950ml", "xlarge": "2000ml", "xxlarge": "4000ml",
}
CATALOG_SUFFIX = {"0125": "125ml", "0250": "250ml", "0500": "500ml",
                  "1000": "950ml", "2000": "2000ml", "4000": "4000ml"}

# How a jar turns into a label
LABEL_HEIGHT_FRACTION = 0.60  # label height = 60% of jar height
MAX_WRAP_FRACTION = 0.85      # label covers at most 85% of the circumference (15% left clear)


def resolve_jar_size(value: str) -> tuple[JarSpec | None, str]:
    """'F', 'size f', '16x10.3cm' (circumference x height), '70x150' (diameter x
    height, mm), '250 mL', 'small', '121-0250' -> JarSpec.  Returns (spec, note)."""
    raw = clean(value)
    if not raw:
        return None, ""
    v = re.sub(r"[\s\-_]|(?<!\d)\.|\.(?!\d)", "", raw.lower())   # keep decimal points
    v = v.replace("ounces", "oz").replace("ounce", "oz")
    letter = re.sub(r"^(size|class|cat|category|jar)", "", v)
    if letter.upper() in JAR_SIZES and len(letter) <= 2:
        return JAR_SIZES[letter.upper()], ""
    m = re.fullmatch(r"(\d+(?:\.\d+)?)[x\u00d7*](\d+(?:\.\d+)?)", v)
    if m and max(float(m.group(1)), float(m.group(2))) <= 60:
        v += "cm"       # bare "13x8.5": circumference x height in cm (jar_sizes.csv units)
    if (m := re.fullmatch(r"(\d+(?:\.\d+)?)[x\u00d7*](\d+(?:\.\d+)?)cm", v)):
        c, h = float(m.group(1)), float(m.group(2))
        return JarSpec(f"{c:g}x{h:g}cm", f"custom jar, circumference {c:g} cm x height {h:g} cm",
                       c, h), ""
    if (m := re.fullmatch(r"(\d+(?:\.\d+)?)[x\u00d7*](\d+(?:\.\d+)?)(mm)?", v)):
        d, h = float(m.group(1)), float(m.group(2))
        return JarSpec(f"{d:g}x{h:g}mm", f"custom jar, diameter {d:g} mm x height {h:g} mm",
                       round(math.pi * d / 10, 1), h / 10), ""
    if (m := re.search(r"[123]21-?(\d{4})", raw)) and m.group(1) in CATALOG_SUFFIX:
        return JAR_SIZES[CATALOG_SUFFIX[m.group(1)]], ""
    if v in JAR_SIZES:
        return JAR_SIZES[v], ""
    if v in JAR_ALIASES:
        key = JAR_ALIASES[v]
        return JAR_SIZES[key], f"jar size '{raw}' is a guess ({key}); record the size letter instead"
    hit = difflib.get_close_matches(v, list(JAR_ALIASES), n=1, cutoff=0.8)
    if hit:
        key = JAR_ALIASES[hit[0]]
        return JAR_SIZES[key], f"jar size '{raw}' read as '{hit[0]}' ({key}), a guess"
    return None, f"jar size '{raw}' not recognised; label sized to content"


# ---- jar_sizes.csv: the label for each measured jar ------------------------
# One row per measured jar: Circumference (cm), Height (cm), Example H#,
# Paper_use_Height (share of the roll width the label takes: "50%", "80%",
# "full"), Orientation (Landscape), Paper_use_length_cm (label length along the
# roll).  So a row reads: a C x H jar gets a label (share x roll width) tall and
# Paper_use_length_cm long.  The script looks a jar up here by its
# circumference and height.
DEFAULT_JAR_TABLE = Path(__file__).with_name("jar_sizes.csv")
MATCH_TOL_CM = 0.15            # a jar within this of a table row uses that row


@dataclass
class TableRow:
    circ: float
    height: float
    label_h_cm: float
    label_len_cm: float
    orientation: str = "Landscape"
    example: str = ""


JAR_TABLE: list[TableRow] = []


def _pct(v: str, roll_cm: float) -> float:
    """'50%' -> 0.5 x roll, 'full' -> roll, '7' or '7 cm' -> 7."""
    t = clean(v).lower().replace("cm", "").strip()
    if t in ("full", "100%", "all"):
        return roll_cm
    if t.endswith("%"):
        return float(t[:-1]) / 100 * roll_cm
    return float(t)


def load_label_table(path: Path, roll_cm: float) -> int:
    """Read jar_sizes.csv (or the same columns in .xlsx) into JAR_TABLE."""
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        ws = openpyxl.load_workbook(path, data_only=True, read_only=True).worksheets[0]
        rows = [[clean(v) for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        with path.open(newline="", encoding="utf-8-sig") as fh:
            rows = [[clean(v) for v in r] for r in csv.reader(fh)]
    hdr = [h.lower() for h in rows[0]]

    def col(*keys, exclude=()):
        for i, h in enumerate(hdr):
            if all(k in h for k in keys) and not any(x in h for x in exclude):
                return i
        return None
    ic, ih = col("circ"), col("height", exclude=("paper", "label"))
    ilh = col("paper", "height") if col("paper", "height") is not None else col("label", "height")
    ill = col("paper", "length") if col("paper", "length") is not None else col("label", "length")
    io, ie = col("orient"), col("example")
    if None in (ic, ih, ilh, ill):
        sys.exit(f"{path}: need columns Circumference (cm), Height (cm), Paper_use_Height and "
                 "Paper_use_length_cm")
    JAR_TABLE.clear()
    for r in rows[1:]:
        if len(r) <= max(ic, ih, ilh, ill) or not r[ic] or not r[ih] or not r[ilh] or not r[ill]:
            continue
        try:
            JAR_TABLE.append(TableRow(float(r[ic]), float(r[ih]), _pct(r[ilh], roll_cm),
                                      float(r[ill]), r[io] if io is not None and r[io] else "Landscape",
                                      r[ie] if ie is not None else ""))
        except ValueError:
            continue
    return len(JAR_TABLE)


@dataclass
class LabelBox:
    """The label a batch of jars gets."""
    h_cm: float
    len_cm: float
    jar: str = ""          # "C 13 x H 8.5 cm", or "" when the label size was given directly
    note: str = ""         # for the review file

    @property
    def key(self) -> str:
        return f"{self.h_cm:g}x{self.len_cm:g}cm"


def _num(v: str) -> float | None:
    m = re.search(r"\d+(?:\.\d+)?", clean(v).replace(",", "."))
    return float(m.group()) if m else None


def label_for_jar(c: float, h: float, roll_cm: float) -> LabelBox:
    """Look a jar (circumference x height, cm) up in JAR_TABLE.  A jar within
    MATCH_TOL_CM of a row uses it; otherwise the biggest row whose jar is no
    bigger than this one (so the label still fits), flagged."""
    jar = f"C {c:g} x H {h:g} cm"
    if not JAR_TABLE:      # no table: the old rule, 60% of height x 85% of circumference
        return LabelBox(min(roll_cm, round(h * LABEL_HEIGHT_FRACTION, 1)),
                        round(c * MAX_WRAP_FRACTION, 1), jar,
                        "no jar_sizes.csv; label = 60% of jar height x 85% of circumference")
    exact = [t for t in JAR_TABLE if abs(t.circ - c) <= MATCH_TOL_CM and abs(t.height - h) <= MATCH_TOL_CM]
    if exact:
        t = min(exact, key=lambda t: abs(t.circ - c) + abs(t.height - h))
        return LabelBox(t.label_h_cm, t.label_len_cm, jar)
    smaller = [t for t in JAR_TABLE if t.circ <= c + MATCH_TOL_CM and t.height <= h + MATCH_TOL_CM]
    if smaller:
        t = max(smaller, key=lambda t: (t.label_len_cm, t.label_h_cm))
        return LabelBox(t.label_h_cm, t.label_len_cm, jar,
                        f"jar {jar} is not in jar_sizes.csv; used the label for the "
                        f"C {t.circ:g} x H {t.height:g} cm jar")
    t = min(JAR_TABLE, key=lambda t: (t.circ, t.height))
    return LabelBox(t.label_h_cm, t.label_len_cm, jar,
                    f"jar {jar} is smaller than every jar in jar_sizes.csv; used the label for "
                    f"C {t.circ:g} x H {t.height:g} cm - check it fits")


def label_box_for(jar: "Jar", forced_jar: JarSpec | None, forced_label: LabelBox | None,
                  roll_cm: float) -> LabelBox | None:
    """Label size for one jar, from (in order) --label-size, --jar-size, the
    sheet's Label height/length columns, Circumference/Height columns, or the
    Jar Size column.  None when nothing gives a size."""
    if forced_label:
        return forced_label
    if forced_jar:
        return label_for_jar(forced_jar.circumference_cm, forced_jar.height_cm, roll_cm)
    lh, ll = _num(jar.label_h), _num(jar.label_len)
    if lh and ll:
        return LabelBox(min(lh, roll_cm), ll, "", "" if lh <= roll_cm else
                        f"label height {lh:g} cm is more than the {roll_cm:g} cm roll")
    c, h = _num(jar.jar_circ), _num(jar.jar_height)
    if c and h:
        return label_for_jar(c, h, roll_cm)
    spec, note = resolve_jar_size(jar.size_raw)
    if spec:
        box = label_for_jar(spec.circumference_cm, spec.height_cm, roll_cm)
        box.note = "; ".join(x for x in (note, box.note) if x)
        return box
    if note:
        jar.problems.append(note)
    return None


def print_label_table(roll_cm: float):
    print(f"{'jar circumference x height':28} label height x length   example")
    for t in sorted(JAR_TABLE, key=lambda t: (t.circ, t.height)):
        print(f"  C {t.circ:5g} x H {t.height:5g} cm       {t.label_h_cm:4g} x {t.label_len_cm:4g} cm"
              f"         {t.example}")


# ==========================================================================
# Drawing
# ==========================================================================
def draw_label(c: canvas.Canvas, jar: Jar, L: Layout, rotate: bool,
               page_h: float | None = None, cut_line: bool = True):
    """page_h = roll width.  When the label is narrower than the roll it sits
    against the top edge and a dashed cut line marks its bottom edge."""
    st = L.style
    page_h = max(page_h or L.height, L.height)
    if rotate:  # page width = roll width, page height = label length
        c.setPageSize((page_h, L.length))
        c.translate(page_h, 0)
        c.rotate(90)
    else:
        c.setPageSize((L.length, page_h))
    if page_h - L.height > 0.1 * cm:
        if cut_line:
            c.saveState()
            c.setLineWidth(0.5)
            c.setDash(3, 3)
            c.line(0, page_h - L.height, L.length, page_h - L.height)
            c.restoreState()
        c.translate(0, page_h - L.height)
    draw_body(c, jar, L)
    c.showPage()


def draw_body(c: canvas.Canvas, jar: Jar, L: Layout):
    """Draw one label with its bottom-left corner at the current origin, upright."""
    st = L.style
    x0, x1 = st.margin_x, L.length - st.margin_x
    if st.valign == "top":
        y = L.height - st.margin_y
    else:
        y = L.height - max(st.margin_y, (L.height - L.content_h) / 2)

    # --- jar size, small, upper-right corner (--show-size)
    if st.show_size and jar.size_text:     # in the top margin, doesn't move the content
        c.setFont(FONT_REG, SIZE_PT)
        c.drawRightString(L.length - 0.1 * cm, L.height - 0.05 * cm - SIZE_PT, jar.size_text)
        y = min(y, L.height - SIZE_PT - 0.15 * cm)

    top = y
    # --- QR code, top right of the header block
    if L.qr_text and L.qr_size:
        draw_qr(c, L.qr_text, x1 - L.qr_size, top - L.qr_size, L.qr_size)
    text_right = x1 - (L.qr_size + QR_GAP if L.qr_size else 0)
    if L.qr_side:                       # QR column on the right: rules and footer stop short of it
        x1 = text_right

    # --- Jar#, centred over the name (as in the template)
    title_w = max(runs_width(r, L.name_pt) for r in L.name_lines)
    if not L.family_own_line:
        title_w = runs_width(L.name_lines[0], L.name_pt) + stringWidth(
            L.family_text, FONT_REG, L.name_pt)
    if L.jar_text:
        js = L.jar_pt
        y -= js
        c.setFont(FONT_REG, js)
        jw = stringWidth(L.jar_text, FONT_REG, js)
        jx = min(max(x0, x0 + (title_w - jw) / 2), max(x0, text_right - jw))
        c.drawString(jx, y, L.jar_text)
        y -= js * 0.25

    # --- name (one or two lines) and family
    ns = L.name_pt
    for k, runs in enumerate(L.name_lines):
        y -= ns
        x = x0
        for text, font in runs:
            text = text.lstrip() if x == x0 else text
            c.setFont(font, ns)
            c.drawString(x, y, text)
            x += stringWidth(text, font, ns)
        if k == 0 and L.family_text and not L.family_own_line:
            c.setFont(FONT_REG, ns)
            c.drawString(x, y, L.family_text)
        y -= ns * 0.25
    if L.family_own_line:
        fs = L.fam_pt
        y -= fs
        c.setFont(FONT_REG, fs)
        c.drawString(x0, y, L.family_text)
        y -= fs * 0.25

    # --- geography
    if jar.header and L.show_geo:
        y -= L.geo_pt
        c.setFont(FONT_REG, L.geo_pt)
        c.drawString(x0, y, jar.header)
        y -= L.geo_pt * 0.3
    y = min(y, top - L.header_h)       # the QR code may be taller than the text
    y -= 4
    c.setLineWidth(st.rule)
    c.line(x0, y, x1, y)

    # --- column headings
    if not st.compact and L.mode == "table":
        y -= st.head_size * 1.15
        c.setFont(FONT_REG, st.head_size)
        for g in range(L.groups):
            cx = x0 + g * (L.group_w + L.group_gap)
            for col in L.columns:
                c.drawString(cx, y, col.title)
                cx += col.width
        y -= st.head_size * 0.45 + 3
        c.setLineWidth(st.thin)
        c.line(x0, y, x1, y)

    # --- rows
    n = len(L.columns[0].values)
    for i in range(n):
        g, r = divmod(i, L.rows_per_group)
        cx = x0 + g * (L.group_w + L.group_gap)
        yy = y - (r + 1) * st.row_pitch + (st.row_pitch - st.row_size) * 0.45
        for col in L.columns:
            c.setFont(FONT_REG, col.size)
            c.drawString(cx, yy, col.values[i])
            cx += col.width
    y -= L.rows_per_group * st.row_pitch + 4
    c.setLineWidth(st.rule)
    c.line(x0, y, x1, y)

    # --- footer, centred, same size as the geography line
    y -= 3
    c.setFont(FONT_REG, st.footer_size)
    for line in L.footer_lines:
        y -= st.footer_size * 1.2
        c.drawCentredString((x0 + x1) / 2, y + st.footer_size * 0.2, line)
    if st.show_preservative and jar.preservative and not st.compact:
        c.setFont(FONT_REG, st.s(8))
        c.drawRightString(x1, y + st.footer_size * 0.2, jar.preservative)


# ==========================================================================
# Long strips: labels end to end down the roll, short ones paired
# ==========================================================================
# With --strip the PDF is one long page per stretch of roll (page width = roll
# width), like a continuous label sheet.  Each label is turned so its length
# runs down the roll.  Two labels whose heights add up to no more than the roll
# width share the same stretch of roll side by side (e.g. two 5 cm labels,
# or a class D with a class L: 4.2 + 5.7 cm on a 10 cm roll).  Which labels to
# pair is chosen to use the least roll: pairing two labels saves the length of
# the shorter one, and the pairs that save the most in total are found with a
# maximum-weight matching.  Dashed lines mark the cuts.
@dataclass
class Slot:
    items: list            # [(jar, Layout)], 1 or 2
    length: float          # along the roll


def pair_labels(placed: list, roll_w: float, allow_pairs: bool = True) -> list[Slot]:
    """placed = [(jar, Layout)] in print order.  Returns slots in print order."""
    n = len(placed)
    partner: dict[int, int] = {}
    fits = lambda i, j: placed[i][1].height + placed[j][1].height <= roll_w + 0.5
    if allow_pairs and n > 1 and (n > 150 or all(fits(i, j) for i in range(n) for j in range(i))):
        # Every label pairs with every other (a one-size batch of short labels), or
        # the batch is too big for the exact matching: longest with next-longest
        # saves the most roll (pairing saves the shorter label's length).
        order = sorted(range(n), key=lambda i: -placed[i][1].length)
        free = list(order)
        while free:
            i = free.pop(0)
            j = next((k for k in free if fits(i, k)), None)
            if j is not None:
                free.remove(j)
                partner[i], partner[j] = j, i
    elif allow_pairs and n > 1:
        import networkx as nx
        G = nx.Graph()
        idx = sorted(range(n), key=lambda i: placed[i][1].height)
        for a_, i in enumerate(idx):
            hi = placed[i][1].height
            if hi * 2 > roll_w + 0.5 and a_ > 0:
                pass
            for j in idx[a_ + 1:]:
                if hi + placed[j][1].height > roll_w + 0.5:
                    break                        # sorted by height: no taller j fits either
                G.add_edge(i, j, weight=min(placed[i][1].length, placed[j][1].length))
        for i, j in nx.max_weight_matching(G):
            partner[i], partner[j] = j, i
    slots, done = [], set()
    for i in range(n):
        if i in done:
            continue
        items = [placed[i]]
        if i in partner:
            j = partner[i]
            items.append(placed[j])
            done.add(j)
        done.add(i)
        slots.append(Slot(items, max(L.length for _, L in items)))
    return slots


def write_strips(c: canvas.Canvas, slots: list[Slot], roll_w: float, gap: float,
                 max_len: float, flip: bool = False) -> list[float]:
    """Draw slots onto long pages; returns the length of each strip (points)."""
    strips: list[list[Slot]] = [[]]
    run = gap
    for sl in slots:
        need = sl.length + gap
        if strips[-1] and run + need > max_len:
            strips.append([])
            run = gap
        strips[-1].append(sl)
        run += need
    lengths = []
    for k, group in enumerate(strips, 1):
        H = gap + sum(sl.length + gap for sl in group)
        lengths.append(H)
        c.setPageSize((roll_w, H))
        y = H - gap
        for pos, sl in enumerate(group, 1):
            lanes = []
            h1 = sl.items[0][1].height
            lanes.append(0.0)
            if len(sl.items) == 2:
                lanes.append(roll_w - sl.items[1][1].height)    # second label against the far edge
            for (jar, L), x in zip(sl.items, lanes):
                jar.position = f"strip {k}, label {pos}" + (
                    "" if len(sl.items) == 1 else (" (left)" if x == 0 else " (right)"))
                jar.strip = k
                c.saveState()
                if flip:
                    c.translate(x + L.height, y - L.length)
                    c.rotate(90)
                else:
                    c.translate(x, y)
                    c.rotate(-90)
                draw_body(c, jar, L)
                c.restoreState()
            # cut marks
            c.saveState()
            c.setLineWidth(0.5)
            c.setDash(3, 3)
            if len(sl.items) == 2:
                h2 = sl.items[1][1].height
                xm = (h1 + roll_w - h2) / 2
                c.line(xm, y, xm, y - sl.length)
            elif h1 < roll_w - 0.1 * cm:
                c.line(h1, y, h1, y - sl.length)
            for (jar, L), x in zip(sl.items, lanes):      # end of a shorter label in a pair
                if L.length < sl.length - 0.1 * cm:
                    c.line(x, y - L.length, x + L.height, y - L.length)
            yb = y - sl.length - gap / 2
            if pos < len(group):
                c.line(0, yb, roll_w, yb)
            c.restoreState()
            y -= sl.length + gap
        c.showPage()
    return lengths


# ==========================================================================
# CLI
# ==========================================================================
def main(argv=None):
    global QR_DATE, QR_COLLECTION, QR_PREFIX
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("xlsx", type=Path, nargs="?", help="jar inventory workbook (.xlsx) or .csv")
    p.add_argument("-o", "--output", type=Path, default=Path("jar_labels.pdf"))
    p.add_argument("--sheet", nargs="+", help="only these sheets (default: every sheet with "
                   "'Catalog #s' and 'Scientific Name' headers)")
    p.add_argument("--jar", nargs="+", help="only these Jar #s, e.g. H1874 H1875")
    p.add_argument("--limit", type=int, help="stop after N labels (test prints)")
    p.add_argument("--paper-width", type=float, default=10.0,
                   help="roll width in cm = label height (default 10)")
    p.add_argument("--length", type=float, help="fixed label length in cm (default: fit content)")
    p.add_argument("--jar-size", metavar="CxH",
                   help="label every jar as this jar: circumference x height in cm (13x8.5), "
                        "looked up in jar_sizes.csv.  Without it each row's Circumference (cm) "
                        "and Height (cm) cells are used (or Label height/length, or Jar Size).")
    p.add_argument("--label-size", metavar="HxL",
                   help="give the label size directly for every label: height x length in cm "
                        "(e.g. 5x12)")
    p.add_argument("--list-jar-sizes", action="store_true", help="show the jar table and exit")
    p.add_argument("--jar-table", type=Path, metavar="CSV", default=DEFAULT_JAR_TABLE,
                   help="jar -> label table (default jar_sizes.csv next to this script): "
                        "Circumference (cm), Height (cm), Paper_use_Height, Paper_use_length_cm")
    p.add_argument("--wrap", type=float, default=MAX_WRAP_FRACTION, help=argparse.SUPPRESS)
    p.add_argument("--min-scale", type=float, default=0.6, help=argparse.SUPPRESS)  # no longer used
    p.add_argument("--min-pt", type=float, default=MIN_PT,
                   help="smallest specimen-row type size in pt (default 10.5 = H1033); the rest "
                        "of the label scales with it")
    p.add_argument("--compact-footer", nargs="+", metavar="LINE",
                   help="footer for compact labels (default: one line, 'UCSB Vertebrate "
                        "Zoology – HERPETOLOGY – CCBER')")
    p.add_argument("--portal", type=Path, metavar="DWCA",
                   help="Symbiota Darwin Core Archive (.zip, folder or occurrences.csv); fills "
                        "each specimen's Locality and Collector from the portal")
    p.add_argument("--reconciliation", type=Path, metavar="XLSX",
                   help="workbook from herp_reconcile.py; only jars whose status is in --status "
                        "are printed")
    p.add_argument("--status", nargs="+", default=["All matched"],
                   help="jar statuses to print with --reconciliation (default: 'All matched')")
    p.add_argument("--exact-names", action="store_true",
                   help="with --reconciliation, also skip jars whose name differs from the "
                        "portal (synonyms, older names, spelling)")
    p.add_argument("--no-cut-line", action="store_true",
                   help="omit the dashed trim line under labels narrower than the roll")
    p.add_argument("--max-length", type=float, default=30.0,
                   help="flag labels longer than this in the review file (cm; default 30)")
    p.add_argument("--scale", type=float, default=1.0, help="scale all type, e.g. 0.9")
    p.add_argument("--valign", choices=["center", "top"], default="center",
                   help="where the label sits across the roll")
    p.add_argument("--footer", nargs="+", metavar="LINE", help="replace the footer lines")
    p.add_argument("--font-regular", help="TrueType file for regular text (e.g. Avenir Book .ttf)")
    p.add_argument("--font-italic", help="TrueType file for the italic scientific name")
    p.add_argument("--no-gbif", action="store_true",
                   help="don't query GBIF for families; use family_cache.json only")
    p.add_argument("--family-cache", type=Path,
                   default=Path(__file__).with_name("family_cache.json"))
    p.add_argument("--no-family", action="store_true", help="leave the family off")
    p.add_argument("--min-run", type=int, default=2,
                   help="consecutive numbers needed before they print as a range "
                        "(default 2: 40748, 40749 -> '40748 - 40749')")
    p.add_argument("--no-ranges", action="store_true",
                   help="print every catalog number on its own row instead of as ranges")
    p.add_argument("--max-expand", type=int, default=300,
                   help="spreadsheet ranges spanning more numbers than this are treated as "
                        "suspect: printed as written and flagged")
    p.add_argument("--preservative", action="store_true", help="print preservative bottom-right")
    p.add_argument("--strip", action="store_true",
                   help="print on long pages (one per stretch of roll), labels end to end down "
                        "the roll, pairing short labels side by side to save paper")
    p.add_argument("--no-pairing", action="store_true", help="with --strip, one label per row")
    p.add_argument("--strip-length", type=float, default=55.9,
                   help="longest page in --strip mode, cm (default 55.9 = 22 in)")
    p.add_argument("--gap", type=float, default=0.3,
                   help="space between labels along the roll in --strip mode, cm (default 0.3)")
    p.add_argument("--flip", action="store_true",
                   help="with --strip, turn labels the other way (text reads up the roll)")
    p.add_argument("--no-qr", action="store_true",
                   help="leave the QR code off the labels")
    p.add_argument("--qr-dots", type=int, choices=[2, 3, 4], default=QR_DOTS,
                   help="QR module size in printer dots (default 3 = 0.375 mm; 2-dot codes "
                        "did not scan reliably in testing)")
    p.add_argument("--qr-date", default=QR_DATE,
                   help="'Updated' date in the QR code (default: today, YYYY-MM-DD)")
    p.add_argument("--qr-collection", default=QR_COLLECTION,
                   help="collection name on the QR code's first line (default 'UCSB Herpetology')")
    p.add_argument("--qr-prefix", default=QR_PREFIX,
                   help="prefix for each catalog number in the QR code (default UCSB-HERP)")
    p.add_argument("--show-size", action="store_true",
                   help="print the jar and label size small in the "
                        "upper-right corner (for test prints)")
    p.add_argument("--rotate", action="store_true",
                   help="rotate pages 90° (for drivers set to portrait)")
    a = p.parse_args(argv)
    if a.jar_table and a.jar_table.exists():
        print(f"Jar table: {load_label_table(a.jar_table, a.paper_width)} jars from {a.jar_table}")
    else:
        print(f"Jar table {a.jar_table} not found; labels = 60% of jar height x 85% of circumference")
    if a.list_jar_sizes:
        print_label_table(a.paper_width)
        return
    if not a.xlsx:
        p.error("the workbook path is required")
    forced_jar = forced_label = None
    if a.jar_size:
        forced_jar, note = resolve_jar_size(a.jar_size)
        if forced_jar is None:
            p.error(f"unknown --jar-size '{a.jar_size}' (use circumference x height in cm, e.g. 13x8.5)")
        if note:
            print(note)
    if a.label_size:
        m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*[x\u00d7*]\s*(\d+(?:\.\d+)?)\s*(cm)?\s*", a.label_size)
        if not m:
            p.error("--label-size is height x length in cm, e.g. 5x12")
        forced_label = LabelBox(min(float(m.group(1)), a.paper_width), float(m.group(2)))

    print("Font:", setup_fonts(a.font_regular, a.font_italic))
    families = None if a.no_family else FamilyLookup(a.family_cache, online=not a.no_gbif)
    jar_filter = {j.upper() for j in a.jar} if a.jar else None
    recon = None
    if a.reconciliation:
        recon = load_reconciliation(a.reconciliation, a.status, a.exact_names)
        print(f"Reconciliation: {sum(map(len, recon.values()))} jars with status "
              f"{' / '.join(a.status)}{' and exact names' if a.exact_names else ''}")
    jars, skipped = read_jars(a.xlsx, a.sheet, jar_filter, a.max_expand, True, families, recon)
    if families:
        families.save()
    if a.limit:
        jars = jars[: a.limit]
    if not jars:
        sys.exit("No jars matched.")

    portal = None
    if a.portal:
        portal = load_portal(a.portal)
        print(f"Portal: {sum(map(len, portal.values())):,} records from {a.portal}")
        for jar in jars:
            attach_portal(jar, portal)

    st = Style(scale=a.scale, valign=a.valign, show_preservative=a.preservative,
               min_run=a.min_run, ranges=not a.no_ranges, min_pt=a.min_pt,
               show_size=a.show_size, qr=not a.no_qr, qr_dots=a.qr_dots)
    QR_DATE, QR_COLLECTION, QR_PREFIX = a.qr_date, a.qr_collection, a.qr_prefix
    if a.footer:
        st.footer = a.footer
    if a.compact_footer:
        st.compact_footer = a.compact_footer
    paper_h = a.paper_width * cm
    fixed_len = a.length * cm if a.length else None

    roll_cm = a.paper_width
    max_len = a.max_length * cm if a.max_length else None
    batches: dict[str, list] = {}            # label size -> [(jar, Layout)], in sheet order
    for jar in jars:
        jar.n_specimens = len(jar.catalogs)
        box = label_box_for(jar, forced_jar, forced_label, roll_cm)
        if box is None:
            jar.problems.append("no jar size: fill in Circumference (cm) and Height (cm), or "
                                "Label height/length, or use --jar-size; label sized to content")
            F = fit_label(jar, st, paper_h, fixed_len or max_len or 30 * cm)
            key, label_desc = "no-size", "no jar size"
        else:
            if box.note:
                jar.problems.append(box.note)
            box_h, box_len = box.h_cm * cm, (fixed_len or box.len_cm * cm)
            jar.size_text = (f"{box.jar} \u2192 " if box.jar else "") + \
                f"label {box.h_cm:g} \u00d7 {box_len / cm:g} cm"
            F = fit_label(jar, st, box_h, box_len)
            key, label_desc = box.key, f"label {box.h_cm:g} x {box_len / cm:g} cm"
        L = F.layout
        jar.fit_note = (f"{label_desc}" + (f" ({box.jar})" if box and box.jar else "")
                        + f": printed {L.height / cm:.1f} x {L.length / cm:.1f} cm, "
                        f"{L.mode} style, rows {F.row_pt:.1f} pt"
                        + (", QR code" if L.qr_text else "")
                        + (", longer localities" if L.level == 2 and jar.specimens else ""))
        jar.level = L.level
        if F.fits and L.mode == "numbers" and jar.specimens:
            jar.problems.append("label too small for locality and collector; catalog numbers only")
        elif F.fits and L.mode == "summary":
            jar.problems.append("catalog numbers don't fit the label; count and range printed, "
                                "full list in the QR code" if L.qr_text else
                                "catalog numbers don't fit the label; count and range only")
        elif F.fits and L.mode == "minimal":
            jar.problems.append("label too small for the full label; Jar#, name, count"
                                + (" and QR code" if L.qr_text else "") + " only "
                                "(no geography/footer)")
        if F.fits and F.row_pt < st.min_pt - 0.05:
            jar.problems.append(f"label too small for H1033-size type; printed at {F.row_pt:.1f} pt")
        if not F.fits:
            jar.problems.append(
                f"doesn't fit the {label_desc} even at the smallest type ({st.min_pt:g} pt); "
                f"printed {L.length / cm:.1f} cm long - use a bigger jar or split the jar")
        if box is None and max_len and L.length > max_len + 0.5:
            jar.problems.append(f"label is {L.length / cm:.0f} cm long (> --max-length "
                                f"{a.max_length:g} cm)")
        batches.setdefault(key, []).append((jar, L))

    notes = []
    for key, placed in batches.items():
        out = a.output if len(batches) == 1 else \
            a.output.with_name(f"{a.output.stem}_{key}{a.output.suffix}")
        out.parent.mkdir(parents=True, exist_ok=True)
        c = canvas.Canvas(str(out))
        c.setTitle(f"Jar labels {key}")
        for jar, _ in placed:
            jar.batch = out.name
        if a.strip:
            slots = pair_labels(placed, paper_h, not a.no_pairing)
            lengths = write_strips(c, slots, paper_h, a.gap * cm, a.strip_length * cm, a.flip)
            single = sum(L.length + a.gap * cm for _, L in placed) + a.gap * cm * len(lengths)
            used = sum(lengths)
            npairs = sum(1 for sl in slots if len(sl.items) == 2)
            roll = (f"{used / cm / 100:.2f} m of roll on {len(lengths)} strip(s)"
                    + (f", {npairs} pairs side by side (saves {1 - used / single:.0%})" if npairs else "")
                    + "\n" + "\n".join(f"      page {k}: {ln / cm:.1f} cm" for k, ln in enumerate(lengths, 1))
                    + f"\n      (no page longer than --strip-length {a.strip_length:g} cm)")
        else:
            for page, (jar, L) in enumerate(placed, 1):
                jar.position = f"page {page}"
                draw_label(c, jar, L, a.rotate, paper_h, not a.no_cut_line)
            roll = f"{sum(L.length for _, L in placed) / cm / 100:.2f} m of roll"
        c.save()
        n_qr = sum(1 for _, L in placed if L.qr_text)
        notes.append(f"  {out.name}: {len(placed)} label(s), {roll}"
                     + (f"; {n_qr} with a QR code" if n_qr else ""))

    review = a.output.with_name(a.output.stem + "_review.csv")
    with review.open("w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["batch", "position", "sheet", "row", "jar", "scientific_name", "n_specimens",
                    "label", "issue"])
        for jar in jars:
            for issue in jar.problems:
                w.writerow([jar.batch, jar.position, jar.sheet, jar.row, jar.jar_id, jar.sciname,
                            jar.n_specimens, jar.fit_note, issue])
        for s in skipped:
            w.writerow(["", "", s["sheet"], s["row"], s["jar"], "", 0, "", s["issue"]])

    if portal is not None:
        spec_csv = a.output.with_name(a.output.stem + "_specimens.csv")
        with spec_csv.open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["batch", "position", "jar", "catalog", "name_on_label", "portal_name", "county",
                        "locality_on_label", "portal_locality",
                        "collector_on_label", "portal_collector", "in_portal"])
            for jar in jars:
                lvl = getattr(jar, "level", 1)
                for s_ in jar.specimens:
                    on_label = (s_.county if jar.loc_mode == "county"
                                else (s_.loc2 or s_.loc1) if lvl >= 2 else s_.loc1)
                    w.writerow([jar.batch, jar.position, jar.jar_id, s_.text, jar.sciname, s_.sciname,
                                s_.county, on_label, s_.locality,
                                s_.collector, s_.collector_raw, "yes" if s_.found else "no"])
        print(f"Specimen localities/collectors -> {spec_csv}")
    print(f"{len(jars)} labels in {len(batches)} batch(es), one PDF per label size:")
    print("\n".join(notes))
    print(f"{sum(1 for j in jars if j.problems)} labels flagged, "
          f"{len(skipped)} rows skipped -> {review}")


if __name__ == "__main__":
    main()

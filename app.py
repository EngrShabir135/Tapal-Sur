"""
Tapal DD 900gm Collectible - Survey Appender (Streamlit)

Upload:
  1) the new Tapal survey export (.csv or .xlsx)
  2) your current  Collectible_Summary_-_BA_Stores.xlsx
  3) your current  DD_900gm_Mug_CP_Survey_*.xlsx
The app APPENDS the new dates / stores into both workbooks (old data of other dates is kept)
and gives you the two updated files to download.

Run:  streamlit run app.py
"""
import io, re, csv, copy, zipfile, datetime, difflib, collections, warnings
import openpyxl
import pandas as pd
from openpyxl.worksheet.dimensions import ColumnDimension

warnings.filterwarnings("ignore")

WINDOW = 25                      # number of date columns in both workbooks
DEFAULT_ALIAS = {                # survey name -> name already used in the workbooks
    "imitaz supermarket (sialkot)": "imtiaz supermarket (sialkot)",
    "time store (bahawalpur)": "auniqe super store",
}
ADD_NEW = "➕ Add as new store"

norm = lambda x: re.sub(r"\s+", " ", str(x)).strip().lower()
isnum = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)
ordn = lambda n: "%d%s" % (n, "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))


# ----------------------------------------------------------------------------- survey parsing
def _to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    s = str(v).strip()
    try: return datetime.date.fromisoformat(s[:10])
    except Exception: return pd.to_datetime(s, dayfirst=False).date()

def _to_time(v):
    if isinstance(v, datetime.datetime): return v.strftime("%H:%M:%S")
    if isinstance(v, datetime.time): return v.strftime("%H:%M:%S")
    p = (str(v).strip().split(" ")[-1]).split(":")
    p = [x.zfill(2) for x in p] + ["00"] * (3 - len(p))
    return ":".join(p[:3])

def _ans(v):
    if v is None: return None
    if isnum(v): return v
    s = str(v).strip()
    if s == "": return None
    try:
        f = float(s.replace(",", ""))
        return int(f) if f == int(f) else f
    except Exception:
        return s

def read_rows(data: bytes, filename: str):
    """returns list of dict rows with keys ObjectID, City, Store, Date, Time, Question, Answer"""
    out = []
    if filename.lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        for ws in wb:
            it = ws.iter_rows(values_only=True)
            hdr = None
            for r in it:
                if hdr is None:
                    if r and "Question" in r and "Answer" in r and "Store" in r: hdr = list(r)
                    continue
                out.append({h: r[i] for i, h in enumerate(hdr) if h})
            if out: break
    else:
        text = data.decode("utf-8-sig", errors="replace")
        rows = list(csv.reader(io.StringIO(text, newline="")))
        hdr = rows[0]
        for r in rows[1:]:
            if not any(str(x).strip() for x in r): continue
            r = list(r) + [""] * (len(hdr) - len(r))
            out.append({h: r[i] for i, h in enumerate(hdr) if h})
    return out

def parse_survey(data: bytes, filename: str):
    subs = {}
    for r in read_rows(data, filename):
        oid, q = r.get("ObjectID"), r.get("Question")
        if not oid or not q or r.get("Store") in (None, ""): continue
        d = subs.setdefault(oid, dict(city=str(r["City"]).strip(), store=str(r["Store"]).strip(),
                                       date=_to_date(r["Date"]), time=_to_time(r.get("Time") or "00:00:00"), ba=r.get("BA")))
        d[str(q).strip()[:2]] = _ans(r.get("Answer"))
    by = collections.OrderedDict()
    for d in subs.values(): by.setdefault(d["store"], []).append(d)
    info = collections.OrderedDict()
    for st, L in by.items():
        L.sort(key=lambda d: (d["date"], d["time"]))
        daily = {}
        for d in L:                                    # last submission of the day wins
            if isnum(d.get("Q2")): daily[d["date"]] = d["Q2"]
            else: daily.setdefault(d["date"], None)
        last = L[-1]; q1, q2 = last.get("Q1"), last.get("Q2")
        if isnum(q1) and q1 > 0: status = "Available"
        elif q1 is None and isnum(q2) and q2 > 0: status = "Available"      # price blank but units sold
        else: status = "Not Available"
        info[norm(st)] = dict(name=st, city=last["city"], daily=daily, status=status,
                              outstock=("out" in str(q1).lower()), lastdate=last["date"])
    return info


# ----------------------------------------------------------------------------- helpers
def table_of(ws):
    return list(ws.tables.values())[0]

def tbl_last_row(t):
    return int(re.findall(r"\d+", t.ref.split(":")[1])[0])

def tfs(obj, fn):
    if obj is not None and getattr(obj, "attr_text", None): obj.attr_text = fn(obj.attr_text)

def parse_header(text, ref_year):
    t = str(text)
    m = re.search(r"\((\d+)(?:st|nd|rd|th) (\w+) (\d{4})\)", t)
    if m: return datetime.datetime.strptime(f"{m.group(1)} {m.group(2)[:3]} {m.group(3)}", "%d %b %Y").date()
    m = re.match(r"\s*(\d+)-(\w{3})", t)
    d = datetime.datetime.strptime(f"{m.group(1)} {m.group(2)} {ref_year}", "%d %b %Y").date()
    return d

def make_header(kind, d):
    return f"{d.day}-{d:%b}" if kind == "cs" else f"Sellout \n({ordn(d.day)} {d:%b} {d.year})"

def replace_names(wb, pairs):
    """bracket-delimited, collision-safe replacement of table column names in every formula"""
    toks = [("[" + a + "]", "\x00%d\x00" % i, "[" + b + "]") for i, (a, b) in enumerate(pairs)]
    def fn(s):
        for a, tk, b in toks: s = s.replace(a, tk)
        for a, tk, b in toks: s = s.replace(tk, b)
        return s
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="): c.value = fn(c.value)
        for t in ws.tables.values():
            for tc in t.tableColumns:
                tfs(tc.calculatedColumnFormula, fn); tfs(tc.totalsRowFormula, fn)
    return fn

def set_dims(ws, c0, c1, first_vis, width):
    """hide the empty leading date columns, show the rest"""
    cd = ws.column_dimensions
    for k in [k for k, v in cd.items() if v.min and v.max and not (v.max < c0 or v.min > c1)]:
        del cd[k]
    L = openpyxl.utils.get_column_letter
    if first_vis > 0:
        cd[L(c0)] = ColumnDimension(ws, index=L(c0), min=c0, max=c0 + first_vis - 1, width=width, hidden=True, customWidth=True)
    cd[L(c0 + first_vis)] = ColumnDimension(ws, index=L(c0 + first_vis), min=c0 + first_vis, max=c1, width=width, hidden=False, customWidth=True)


def guess_channel(name):
    n = name.lower()
    return "IMT" if ("imtiaz" in n or n.startswith("spar") or "carrefour" in n or "metro" in n or n.startswith("csd")) else "LMT"

CITY_FALLBACK = {  # city -> (region, zone) for cities that may not exist in the workbooks yet
    "quetta": ("Sukkur", "South"), "muzaffarabad": ("Islamabad", "North"), "rawalpindi": ("Islamabad", "North"),
    "daska": ("Gujranwala", "North"), "wazirabad": ("Gujranwala", "North"), "sahiwal": ("Faisalabad", "North"),
    "rahim yar khan": ("Bahawalpur", "North"), "peshawar": ("Islamabad", "North"), "sialkot": ("Gujranwala", "North"),
    "sargodha": ("Faisalabad", "North"), "jhelum": ("Gujranwala", "North"), "gujrat": ("Gujranwala", "North"),
    "hyderabad": ("Hyderabad", "South"), "karachi": ("Karachi", "South"), "lahore": ("Lahore", "North"),
    "multan": ("Multan", "North"), "faisalabad": ("Faisalabad", "North"), "islamabad": ("Islamabad", "North"),
}


# ----------------------------------------------------------------------------- workbook context
def load_ctx(cs_bytes, dd_bytes):
    cwb = openpyxl.load_workbook(io.BytesIO(cs_bytes))
    dwb = openpyxl.load_workbook(io.BytesIO(dd_bytes))
    return cwb, dwb

def existing_stores(cwb, dwb):
    s = {}
    for n in ("GT", "KA"):
        ws = cwb[n]
        for r in range(2, tbl_last_row(table_of(ws)) + 1):
            if ws.cell(r, 3).value: s[norm(ws.cell(r, 3).value)] = ws.cell(r, 3).value
    ws = dwb["Compile"]; t = table_of(ws)
    for r in range(7, tbl_last_row(t)):
        if ws.cell(r, 4).value: s.setdefault(norm(ws.cell(r, 4).value), ws.cell(r, 4).value)
    return s

def city_map(cwb, dwb):
    m = collections.defaultdict(collections.Counter)
    for n in ("GT", "KA"):
        ws = cwb[n]
        for r in range(2, tbl_last_row(table_of(ws)) + 1):
            if ws.cell(r, 3).value: m[norm(ws.cell(r, 2).value)][(ws.cell(r, 1).value, ws.cell(r, 4).value, ws.cell(r, 2).value)] += 1
    return m

def suggest_new(info, existing, cwb, dwb, alias=DEFAULT_ALIAS):
    """-> (matched {survey_key: file_key}, review dataframe for unmatched stores)"""
    matched, rows = {}, []
    cm = city_map(cwb, dwb)
    regions = sorted({k[0] for c in cm.values() for k in c if k[0]})
    for k, v in info.items():
        t = alias.get(k, k)
        if t in existing: matched[k] = t; continue
        c = v["city"].replace("_", " "); ck = norm(c)
        if ck not in cm:
            near = difflib.get_close_matches(ck, list(cm), 1, 0.85)
            if near: ck = near[0]
        if ck in cm:
            reg, zone, city = cm[ck].most_common(1)[0][0]
        else:
            reg, zone = CITY_FALLBACK.get(ck, ("", "")); city = c
        close = difflib.get_close_matches(k, list(existing), 1, 0.8)
        rows.append({"Survey store": v["name"], "Survey city": c, "Possible existing match": existing[close[0]] if close else "",
                     "Action": ADD_NEW, "City": city, "Region": reg, "Zone": zone, "Channel": guess_channel(v["name"])})
    return matched, pd.DataFrame(rows), regions


# ----------------------------------------------------------------------------- core update
def roll_window(ws, kind, spec, new_end, ref_year, report):
    """make the 25 date columns end at new_end (shifting old values left if needed). returns new date list"""
    t = table_of(ws); c0 = spec["date0"]
    old_hdr = [ws.cell(spec["hdr"], c0 + i).value for i in range(WINDOW)]
    old_end = parse_header(old_hdr[-1], ref_year)
    if kind == "cs" and old_end > new_end and (old_end - new_end).days > 200: old_end = old_end.replace(year=old_end.year - 1)
    new_end = max(new_end, old_end)
    k = (new_end - old_end).days
    dates = [new_end - datetime.timedelta(days=WINDOW - 1 - i) for i in range(WINDOW)]
    new_hdr = [make_header(kind, d) for d in dates]
    if k > 0:
        last = spec["last_data_row"](ws, t)
        dropped = 0
        for r in range(spec["first"], last + 1):
            vals = [ws.cell(r, c0 + i).value for i in range(WINDOW)]
            dropped += sum(1 for v in vals[:min(k, WINDOW)] if v is not None)
            nv = [vals[i + k] if i + k < WINDOW else None for i in range(WINDOW)]
            for i in range(WINDOW): ws.cell(r, c0 + i).value = nv[i]
        if dropped: report["warnings"].append(f"{ws.title}: {dropped} older sellout values fell out of the {WINDOW}-day window and were dropped.")
    for i in range(WINDOW):
        ws.cell(spec["hdr"], c0 + i).value = new_hdr[i]
        t.tableColumns[c0 - 1 + i].name = new_hdr[i]
    return dates, list(zip(old_hdr, new_hdr))

def sum_row(ws, r, c0):
    return sum(v for v in (ws.cell(r, c0 + i).value for i in range(WINDOW)) if isnum(v))

def apply_row(ws, r, v, spec, dcol):
    c0 = spec["date0"]
    ws.cell(r, spec["status"]).value = v["status"]
    for d, val in v["daily"].items():
        if d in dcol and val is not None: ws.cell(r, c0 + dcol[d]).value = val
    tot = sum_row(ws, r, c0)
    ws.cell(r, spec["rem"]).value = "Stock finished" if v["status"] == "Not Available" and (tot > 0 or v["outstock"]) else None

def first_visible(ws, spec, rows):
    c0 = spec["date0"]
    for i in range(WINDOW):
        if any(ws.cell(r, c0 + i).value is not None for r in rows): return i
    return 0

def process(cs_bytes, dd_bytes, info, matched, review_df):
    report = dict(warnings=[], sheets={})
    cwb, dwb = load_ctx(cs_bytes, dd_bytes)
    max_date = max(d for v in info.values() for d in v["daily"])
    # survey store -> file key (user decisions override)
    target = dict(matched); newrows = []
    ex = existing_stores(cwb, dwb)
    ex_by_name = {v: k for k, v in ex.items()}
    for _, r in review_df.iterrows():
        k = norm(r["Survey store"])
        if r["Action"] != ADD_NEW and r["Action"] in ex_by_name: target[k] = ex_by_name[r["Action"]]
        else:
            if not r["Region"] or not r["Zone"] or r["Channel"] not in ("LMT", "IMT"):
                raise ValueError(f"Please complete Region / Zone / Channel for new store: {r['Survey store']}")
            newrows.append(r)
    byfile = {}
    for sk, fk in target.items(): byfile[fk] = info[sk]
    # ---- Collectible ----
    cs_spec = dict(date0=7, hdr=1, first=2, status=6, rem=33, last_data_row=lambda ws, t: tbl_last_row(t))
    all_pairs_cs = []
    for sheet, chan in (("GT", "LMT"), ("KA", "IMT")):
        ws = cwb[sheet]; t = table_of(ws)
        dates, pairs = roll_window(ws, "cs", cs_spec, max_date, max_date.year, report)
        replace_names(cwb, [p for p in pairs if p[0] != p[1]])
        dcol = {d: i for i, d in enumerate(dates)}
        for v in info.values():
            for d in v["daily"]:
                if d < dates[0]: report["warnings"].append(f"{v['name']}: {d} is older than the {WINDOW}-day window - skipped.")
        maxr = tbl_last_row(t); upd = 0
        for r in range(2, maxr + 1):
            n = ws.cell(r, 3).value
            if n and norm(n) in byfile: apply_row(ws, r, byfile[norm(n)], cs_spec, dcol); upd += 1
        nxt = max(r for r in range(2, maxr + 1) if ws.cell(r, 3).value) + 1; add = 0
        regions_in_summary = {str(ws.cell(r, 35).value).strip() for r in range(1, 40) if ws.cell(r, 35).value}
        for r_ in newrows:
            if r_["Channel"] != chan: continue
            if nxt > maxr:                                    # extend the table by one row
                for c in range(1, 34): ws.cell(maxr + 1, c)._style = copy.copy(ws.cell(maxr, c)._style)
                ws.cell(maxr + 1, 32).value = ws.cell(maxr, 32).value
                for row in ws.iter_rows(min_col=35, max_col=39):
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith("="): c.value = c.value.replace(f"${maxr}", f"${maxr + 1}")
                maxr += 1; t.ref = re.sub(r"\d+$", str(maxr), t.ref)
            ws.cell(nxt, 1).value = r_["Region"]; ws.cell(nxt, 2).value = r_["City"]; ws.cell(nxt, 3).value = r_["Survey store"]
            ws.cell(nxt, 4).value = r_["Zone"]; ws.cell(nxt, 5).value = chan
            apply_row(ws, nxt, info[norm(r_["Survey store"])], cs_spec, dcol)
            if r_["Region"] not in regions_in_summary:
                report["warnings"].append(f"{sheet}: region '{r_['Region']}' is not in the summary block, so it is not counted there.")
            nxt += 1; add += 1
        for row in ws.iter_rows(min_col=35, max_col=39):
            for c in row:
                if isinstance(c.value, str) and "Data is updated" in c.value:
                    c.value = re.sub(r"Data is updated[^\n]*?\.", f"Data is updated till {max_date.day} {max_date:%b %Y}.", c.value, 1)
        rows = [r for r in range(2, nxt)]
        fv = first_visible(ws, cs_spec, rows)
        set_dims(ws, 7, 31, fv, 9.44)
        report["sheets"][sheet] = dict(updated=upd, added=add, window=f"{dates[0]:%d %b} - {dates[-1]:%d %b %Y}")
    cwb.calculation.fullCalcOnLoad = True
    # ---- DD ----
    ws = dwb["Compile"]; t = table_of(ws); tname = t.name
    dd_spec = dict(date0=8, hdr=6, first=7, status=7, rem=34,
                   last_data_row=lambda ws, t: tbl_last_row(t) - 1)
    old_status = ws["G6"].value
    dates, pairs = roll_window(ws, "dd", dd_spec, max_date, max_date.year, report)
    new_status = f"DD 900gm CP Stock Status\n({ordn(max_date.day)} {max_date:%b %Y})"
    ws["G6"].value = new_status; t.tableColumns[6].name = new_status
    ws["AH6"].value = f"Remarks ({max_date:%d/%m/%Y})"
    replace_names(dwb, [p for p in pairs if p[0] != p[1]] + [(old_status, new_status)])
    dcol = {d: i for i, d in enumerate(dates)}
    tot = tbl_last_row(t); last = tot - 1; NC = 34
    upd = 0
    for r in range(7, last + 1):
        n = ws.cell(r, 4).value
        if n and norm(n) in byfile: apply_row(ws, r, byfile[norm(n)], dd_spec, dcol); upd += 1
    add = len(newrows)
    if add:
        totrow = [(ws.cell(tot, c).value, copy.copy(ws.cell(tot, c)._style)) for c in range(1, NC + 1)]
        tpl = [copy.copy(ws.cell(last, c)._style) for c in range(1, NC + 1)]
        th, tp = ws.row_dimensions[tot].height, ws.row_dimensions[last].height
        for i, r_ in enumerate(newrows):
            rr = tot + i
            for c in range(1, NC + 1): ws.cell(rr, c)._style = copy.copy(tpl[c - 1]); ws.cell(rr, c).value = None
            ws.row_dimensions[rr].height = tp
            ws.cell(rr, 2).value = r_["Region"]; ws.cell(rr, 3).value = r_["City"]; ws.cell(rr, 4).value = r_["Survey store"]
            ws.cell(rr, 5).value = r_["Zone"]; ws.cell(rr, 6).value = r_["Channel"]
            apply_row(ws, rr, info[norm(r_["Survey store"])], dd_spec, dcol)
            ws.cell(rr, 33).value = ws.cell(last, 33).value
        newtot, newlast = tot + add, tot + add - 1
        for c in range(1, NC + 1):
            val, sty = totrow[c - 1]; cell = ws.cell(newtot, c); cell._style = sty
            cell.value = val
        ws.row_dimensions[newtot].height = th
        last, tot = newlast, newtot
    for c in range(1, NC + 1):                                 # (re)write total-row ranges
        v = ws.cell(tot, c).value
        if isinstance(v, str) and v.startswith("="):
            ws.cell(tot, c).value = re.sub(r"([A-Z]{1,2})7:([A-Z]{1,2})\d+", lambda m: f"{m.group(1)}7:{m.group(2)}{last}", v)
    t.ref = re.sub(r"\d+$", str(tot), t.ref)
    if t.autoFilter is not None: t.autoFilter.ref = re.sub(r"\d+$", str(last), t.autoFilter.ref)
    t.sortState = None
    # keep the table filter view (e.g. Channel = LMT) consistent for the new rows
    fvals = None
    if t.autoFilter is not None:
        for fc in t.autoFilter.filterColumn:
            if fc.colId == 5 and fc.filters is not None: fvals = list(fc.filters.filter)
    for r in range(7, last + 1):
        ws.row_dimensions[r].hidden = bool(fvals) and ws.cell(r, 6).value not in fvals
    fv = first_visible(ws, dd_spec, range(7, last + 1))
    set_dims(ws, 8, 32, fv, 10.44)
    dwb.calculation.fullCalcOnLoad = True
    report["sheets"]["DD Compile"] = dict(updated=upd, added=add, window=f"{dates[0]:%d %b} - {dates[-1]:%d %b %Y}")
    # ---- save + restore DD slicers ----
    b1 = io.BytesIO(); cwb.save(b1)
    b2 = io.BytesIO(); dwb.save(b2)
    dd_out = restore_slicers(dd_bytes, b2.getvalue())
    name_dd = f"DD_900gm_Mug_CP_Survey__{ordn(max_date.day)}_{max_date:%b}_{max_date.year}_.xlsx"
    return b1.getvalue(), dd_out, name_dd, report


# ----------------------------------------------------------------------------- keep slicers (openpyxl drops them)
def restore_slicers(src_bytes, staged_bytes):
    zs = zipfile.ZipFile(io.BytesIO(src_bytes))
    names = zs.namelist()
    parts = [n for n in names if n.startswith(("xl/slicers/", "xl/slicerCaches/", "xl/drawings/")) and not n.endswith("/")]
    if not any(n.startswith("xl/slicers/") for n in parts): return staged_bytes
    st = zipfile.ZipFile(io.BytesIO(staged_bytes)); files = {n: st.read(n) for n in st.namelist()}
    # which sheet holds the slicers in the source
    sheet_file = None
    for n in names:
        if n.startswith("xl/worksheets/_rels/") and b"/slicer" in zs.read(n) or (n.startswith("xl/worksheets/_rels/") and b"slicers/" in zs.read(n)):
            sheet_file = n.split("/")[-1].replace(".rels", ""); break
    if not sheet_file: return staged_bytes
    for n in parts: files[n] = zs.read(n)
    ctypes = {"xl/drawings/": "application/vnd.openxmlformats-officedocument.drawing+xml",
              "xl/slicers/": "application/vnd.ms-excel.slicer+xml", "xl/slicerCaches/": "application/vnd.ms-excel.slicerCache+xml"}
    ct = files["[Content_Types].xml"].decode(); add = ""
    for n in parts:
        if n.endswith(".rels"): continue
        for p, typ in ctypes.items():
            if n.startswith(p) and f'PartName="/{n}"' not in ct: add += f'<Override PartName="/{n}" ContentType="{typ}"/>'
    files["[Content_Types].xml"] = ct.replace("</Types>", add + "</Types>").encode()
    # sheet rels / xml
    rel_p = f"xl/worksheets/_rels/{sheet_file}.rels"
    src_rel = zs.read(rel_p).decode()
    drawing_t = re.search(r'Target="([^"]*drawings/[^"]+)"', src_rel).group(1).split("drawings/")[-1]
    slicer_t = re.search(r'Target="([^"]*slicers/[^"]+)"', src_rel).group(1).split("slicers/")[-1]
    r = files[rel_p].decode()
    r = r.replace("</Relationships>", f'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="/xl/drawings/{drawing_t}" Id="rId90"/>'
                  f'<Relationship Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="/xl/slicers/{slicer_t}" Id="rId91"/></Relationships>')
    files[rel_p] = r.encode()
    sp = f"xl/worksheets/{sheet_file}"; s = files[sp].decode()
    root = re.match(r"<worksheet[^>]*>", s).group(0)
    if "xmlns:r=" not in root:
        s = s.replace(root, root[:-1] + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">', 1)
    s = s.replace("<tableParts", '<drawing r:id="rId90"/><tableParts', 1)
    s = s.replace("</worksheet>", '<extLst><ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
                  '<x14:slicerList xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicer r:id="rId91"/></x14:slicerList></ext></extLst></worksheet>')
    files[sp] = s.encode()
    # workbook level slicer caches
    wrel = zs.read("xl/_rels/workbook.xml.rels").decode()
    caches = [m for m in re.findall(r'<Relationship [^>]*>', wrel) if "slicerCache" in m]
    wr = files["xl/_rels/workbook.xml.rels"].decode(); items = ""; k = 90
    for m in caches:
        tgt = re.search(r'Target="([^"]+)"', m).group(1).split("slicerCaches/")[-1]
        wr = wr.replace("</Relationships>", f'<Relationship Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="/xl/slicerCaches/{tgt}" Id="rId{k}"/></Relationships>')
        items += f'<x14:slicerCache r:id="rId{k}"/>'; k += 1
    files["xl/_rels/workbook.xml.rels"] = wr.encode()
    w = files["xl/workbook.xml"].decode()
    w = w.replace("</workbook>", '<extLst><ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
                  '<x15:slicerCaches xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                  + items + "</x15:slicerCaches></ext></extLst></workbook>")
    files["xl/workbook.xml"] = w.encode()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", files["[Content_Types].xml"])
        for n, b in files.items():
            if n != "[Content_Types].xml": z.writestr(n, b)
    return out.getvalue()


# ----------------------------------------------------------------------------- Streamlit UI
def main():
    import streamlit as st
    st.set_page_config(page_title="Tapal DD 900gm - Survey Appender", page_icon="📊", layout="wide")
    st.title("📊 Tapal DD 900gm Collectible – Survey Appender")
    st.caption("Upload the new survey export and your two current workbooks. New dates and stores are appended; "
               "everything else in the files is kept.")
    c1, c2, c3 = st.columns(3)
    sv = c1.file_uploader("1️⃣ New survey (CSV or XLSX)", type=["csv", "xlsx"])
    cs = c2.file_uploader("2️⃣ Collectible Summary - BA Stores (xlsx)", type=["xlsx"])
    dd = c3.file_uploader("3️⃣ DD 900gm Mug CP Survey (xlsx)", type=["xlsx"])
    if not (sv and cs and dd):
        st.info("Upload all three files to continue."); return
    sv_b, cs_b, dd_b = sv.getvalue(), cs.getvalue(), dd.getvalue()
    try:
        info = parse_survey(sv_b, sv.name)
    except Exception as e:
        st.error(f"Could not read the survey file: {e}"); return
    if not info: st.error("No survey rows found."); return
    dts = sorted({d for v in info.values() for d in v["daily"]})
    m1, m2, m3 = st.columns(3)
    m1.metric("Stores in survey", len(info)); m2.metric("Dates", len(dts)); m3.metric("Last date", f"{dts[-1]:%d %b %Y}")
    st.write("Dates found: " + ", ".join(f"{d:%d %b}" for d in dts))
    cwb, dwb = load_ctx(cs_b, dd_b)
    ex = existing_stores(cwb, dwb)
    matched, review, regions = suggest_new(info, ex, cwb, dwb)
    st.success(f"{len(matched)} survey stores matched to stores already in your files.")
    if len(review):
        st.subheader("Stores not found in your files")
        st.caption("Keep **Add as new store** to append them, or pick the existing store they correspond to. "
                   "Check Region / Zone / Channel (LMT = old GT sheet, IMT = KA sheet).")
        options = [ADD_NEW] + sorted(ex.values(), key=str.lower)
        review = st.data_editor(review, hide_index=True, use_container_width=True, key="rev", disabled=["Survey store", "Survey city", "Possible existing match"],
                                column_config={"Action": st.column_config.SelectboxColumn(options=options, required=True),
                                               "Region": st.column_config.SelectboxColumn(options=[""] + regions),
                                               "Zone": st.column_config.SelectboxColumn(options=["", "North", "South"]),
                                               "Channel": st.column_config.SelectboxColumn(options=["LMT", "IMT"])})
    if st.button("▶ Append data and build files", type="primary"):
        try:
            with st.spinner("Updating workbooks..."):
                cs_out, dd_out, dd_name, rep = process(cs_b, dd_b, info, matched, review)
            st.session_state["out"] = (cs_out, dd_out, dd_name, rep)
        except Exception as e:
            st.error(str(e)); st.session_state.pop("out", None)
    if "out" in st.session_state:
        cs_out, dd_out, dd_name, rep = st.session_state["out"]
        st.subheader("Done")
        st.dataframe(pd.DataFrame(rep["sheets"]).T.rename(columns={"updated": "Stores updated", "added": "New stores added", "window": "Date window"}))
        for w in rep["warnings"]: st.warning(w)
        d1, d2 = st.columns(2)
        d1.download_button("⬇ Collectible_Summary_-_BA_Stores.xlsx", cs_out, "Collectible_Summary_-_BA_Stores.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button(f"⬇ {dd_name}", dd_out, dd_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Open in Excel – formulas recalculate on open.")

if __name__ == "__main__":
    main()

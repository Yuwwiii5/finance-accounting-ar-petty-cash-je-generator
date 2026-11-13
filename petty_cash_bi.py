"""
Petty cash -> Business Importer upload builder.

Reads a Mews "Payment report" export, keeps the cash payments, and writes them
into the matching property tab of the Petty Cash BI Upload template.

Usage:
    py petty_cash_bi.py                          (opens the window)
    py petty_cash_bi.py --doc 18193 --date 9/15 --counted 63888   (no window)

Needs: pip install openpyxl pandas
"""
import argparse
import datetime as dt
import re
import subprocess
import sys
import zipfile
from copy import copy
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill

HERE = Path(__file__).resolve().parent          # _system: script, template, saved answers
EVIDENCE = HERE.parent if HERE.name.lower() == "_system" else HERE   # reports in, uploads out
TEMPLATE = str(HERE / "Petty_Cash_BI_Upload_Template_fixed.xlsx")

# A copy of every upload file goes to:
#   DOC_ROOT \ <journal month, e.g. 2026.10> \ Petty Cash \ <hotel folder>
# Missing folders are created; an existing hotel folder is reused even if its
# number prefix differs (e.g. "Kabin Machi ARD" for "07. KABIN Machi ARD").
DOC_ROOT = r"C:\path\to\Documentation"  # set to your shared Documentation folder
DOC_FOLDERS = {  # template tab -> hotel folder (same names as Bank Transactions)
    "WHKSBI": "01. Wayfarer Kyoto Shijo ARD",
    "ORI-KBI": "02. ORI Kyoto ARD",
    "ORI-RBI": "03. ORI Rokkaku ARD",
    "KKBI": "04. KABIN Kyoto ARD",
    "KKojiBI": "05. KABIN Koji ARD",
    "KTakaBI": "06. KABIN Taka ARD",
    "KMachiBI": "07. KABIN Machi ARD",
    "KMnmzBI": "08. KABIN Minamiza ARD",
    "WFMatsuBI": "09. Wayfarer Matsu ARD",
    "WFSSBBI": "10. Wayfarer Shinsaibashi ARD",
    "WFShinBI": "11. Wayfarer Shin ARD",
    "WFGojoBI": "12. Wayfarer Gojo ARD",
    "ORI-IBI": "14. ORI Ikebukuro ARD",
}

# Mews "Enterprise" name -> template tab, for names that don't match the tab's
# location (cell J2). Anything not listed here is matched to J2 automatically.
ENTERPRISE_TO_TAB = {
    "KABIN Kyoto": "KKBI",
    "KABIN Taka": "KTakaBI",
    "KABIN Gion": "KMnmzBI",  # Mews calls Kabin Minamiza "KABIN Gion"
}
SKIP_TABS = {"SOP", "Recon Disc"}

YELLOW = PatternFill("solid", start_color="FFFF00")


def ask(prompt, default=None):
    suffix = f" [{default}]" if default else ""
    try:
        answer = input(f"{prompt}{suffix}: ").strip()
    except EOFError:  # nobody to answer (e.g. the app's Run button): take the default
        print(default or "")
        return default
    return answer or default


# The AR master list of hotels (same file the OTA JE tool is based on):
# MEWS Enterprise -> Line: Location (what the template tabs use in J2) and Evidence Naming.
MASTER = r"C:\path\to\Master Query Parameters.xlsx"  # set to your hotel master list
_master = None


def master_row(enterprise):
    """The master's row for a Mews enterprise (dict), or {} if not listed / file not reachable."""
    global _master
    if _master is None:
        try:
            df = pd.read_excel(MASTER, dtype=str)
            _master = {str(r["MEWS Enterprise"]).strip().lower(): r.to_dict() for _, r in df.iterrows()}
        except Exception:
            _master = {}
    return _master.get(enterprise.strip().lower(), {})


def find_tab(enterprise, template=TEMPLATE):
    """Template tab for a Mews enterprise: the mapping first, then cell J2 against the
    master's Line: Location (e.g. WAYFARER Gojo -> Gojo), then J2 against the Mews name."""
    wb = openpyxl.load_workbook(template, read_only=True)
    names = {n.strip(): n for n in wb.sheetnames if n.strip() not in SKIP_TABS}
    if enterprise in ENTERPRISE_TO_TAB:
        return names.get(ENTERPRISE_TO_TAB[enterprise])
    wanted = {enterprise.lower(), str(master_row(enterprise).get("Line: Location", "")).strip().lower()}
    for clean, real in names.items():
        if str(wb[real]["J2"].value or "").strip().lower() in wanted - {""}:
            return real
    return None


def resolve_tab(tab, template=TEMPLATE):
    """Match a typed tab name, ignoring case and stray spaces."""
    sheetnames = openpyxl.load_workbook(template, read_only=True).sheetnames
    for n in sheetnames:
        if n.strip().lower() == tab.strip().lower():
            return n
    raise SystemExit(f"Tab '{tab}' not found. Tabs: {', '.join(n.strip() for n in sheetnames)}")


INLINE_STR = re.compile(r'<c ([^>]*?)t="inlineStr"([^>]*)><is>(?:<t(?: [^>]*)?>(.*?)</t>|<t ?/>)</is></c>',
                        re.DOTALL)


def use_shared_strings(path):
    """Rewrite openpyxl's inline text cells into Excel's normal shared-strings list.

    openpyxl 3.1.3+ stores text inline in each cell. Excel reads that fine, but
    Business Importer appears not to, so every Account and Credit/Debit cell
    looked blank and each line was rejected as a one-line journal entry.
    """
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    strings, index = [], {}

    def to_shared(m):
        text = m.group(3) or ""
        if text not in index:
            index[text] = len(strings)
            strings.append(text)
        return f'<c {m.group(1)}t="s"{m.group(2)}><v>{index[text]}</v></c>'

    for name in [n for n in files if n.startswith("xl/worksheets/sheet")]:
        files[name] = INLINE_STR.sub(to_shared, files[name].decode("utf-8")).encode("utf-8")
    if not strings:
        return
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    files["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<sst xmlns="{ns}" count="{len(strings)}" uniqueCount="{len(strings)}">'
        + "".join(f'<si><t xml:space="preserve">{s}</t></si>' for s in strings)
        + "</sst>").encode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    if "sharedStrings" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/sharedStrings.xml" ContentType='
                        '"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
        files["[Content_Types].xml"] = ct.encode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "sharedStrings" not in rels:
        rels = rels.replace("</Relationships>", '<Relationship Id="rIdSST" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
        files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)


def excel_resave(path):
    """Open and save the file in Excel so it matches a hand-made upload exactly.

    Returns False if Excel isn't available (the shared-strings fix is used instead).
    """
    full = str(Path(path).resolve()).replace("'", "''")
    ps = ("$x = New-Object -ComObject Excel.Application; $x.DisplayAlerts = $false; "
          f"$wb = $x.Workbooks.Open('{full}'); $x.CalculateFull(); $wb.Save(); $wb.Close($false); "
          "$x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)")
    try:
        done = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                              capture_output=True, text=True, timeout=180)
        return done.returncode == 0 and not done.stderr.strip()
    except (OSError, subprocess.TimeoutExpired):
        return False


def remove_old_uploads(keep):
    """Send every other upload file in the folder to the Recycle Bin.

    Recycle Bin rather than a hard delete, so a file removed by mistake can be restored.
    A file that's open in Excel is left alone.
    """
    keep = Path(keep).resolve()
    old = list(keep.parent.glob("*_BI_Upload_Doc*.xlsx")) + list(keep.parent.glob(f"*{UPLOAD_TAG}_*.xlsx"))
    return [f.name for f in old
            if f.resolve() != keep and not f.name.startswith("~$") and to_recycle_bin(f)]


def to_recycle_bin(f):
    """Move one file to the Recycle Bin (recoverable). True if it's gone."""
    f = Path(f)
    path = str(f.resolve()).replace("'", "''")
    ps = ("Add-Type -AssemblyName Microsoft.VisualBasic; "
          "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile("
          f"'{path}', 'OnlyErrorDialogs', 'SendToRecycleBin')")
    done = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                          capture_output=True, text=True)
    if done.returncode == 0 and not f.exists():
        return True
    print(f"Couldn't remove {f.name} (is it open in Excel?)")
    return False


def remove_used_reports(payment_report, enterprise):
    """After a successful upload: the Payment report goes to the Recycle Bin, and the
    hotel's Reservation report too unless another Payment report for that hotel is
    still waiting in the folder (it will need it)."""
    payment_report = Path(payment_report).resolve()
    used = [payment_report]
    others = [f for f in payment_report.parent.glob("Payment*report*.xlsx")
              if f.resolve() != payment_report and not f.name.startswith("~$")]
    still_needed = any(read_cash(f, quiet=True)[0] == enterprise for f in others)
    res = find_reservation_report(payment_report, enterprise)
    if res and not still_needed:
        used.append(res)
    return [f.name for f in used if to_recycle_bin(f)]


def recon_account(wb, ws):
    """Discrepancy account used for this property: the tab itself, then the Recon Disc tab."""
    loc = str(ws["J2"].value or "").strip().lower()
    sheets = [ws] + ([wb["Recon Disc"]] if "Recon Disc" in wb.sheetnames else [])
    for sheet in sheets:
        for r in range(2, sheet.max_row + 1):
            acct = str(sheet.cell(r, 2).value or "")
            if "recon" in acct.lower() and (sheet is ws or
                    str(sheet.cell(r, 10).value or "").strip().lower() == loc):
                return acct
    # No example in the template: use the brand's usual account.
    if loc.startswith("kabin"):
        return "XXXXXXX Reconciliation Discrepancies"
    return "XXXXXXX Recon Dis. / Round. Difference"


UPLOAD_TAG = "_Accommodation_PettyCash"

# Mews "Travel agency" -> customer prefix in QuickBooks, where the spelling differs.
# Anything not listed is used as Mews spells it; a blank travel agency means Direct.
OTA_NAMES = {
    "Hotels.com": "Expedia",  # Hotels.com is part of Expedia
}
DIRECT = "Direct"


def find_reservation_report(payment_report, enterprise):
    """Newest Reservation report in the folder for the same hotel, or None."""
    folder = Path(payment_report).resolve().parent
    for f in sorted(folder.glob("Reservation*report*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True):
        if f.name.startswith("~$"):
            continue
        params = pd.read_excel(f, sheet_name="Parameters", header=None)
        if str(params.loc[params[0] == "Enterprise", 1].iloc[0]).strip() == enterprise:
            return f
    return None


def ota_lookup(reservation_report):
    """Function: guest name -> OTA prefix, or None when unknown or ambiguous."""
    if reservation_report is None:
        print("WARNING: no Reservation report for this hotel in the folder, so no OTA names.\n"
              "         Export it from Mews (same hotel and period) and run again.")
        return lambda guest: None
    res = pd.read_excel(reservation_report, sheet_name="Reservations")
    first = res["First name"].fillna("").astype(str).str.strip()
    last = res["Last name"].fillna("").astype(str).str.strip()
    agency = res["Travel agency"].apply(
        lambda a: DIRECT if pd.isna(a) or not str(a).strip() else OTA_NAMES.get(str(a).strip(), str(a).strip()))
    found = {}
    for f, l, a in zip(first, last, agency):
        for key in {f"{f} {l}".strip().lower(), f"{l} {f}".strip().lower()}:
            if key:
                found.setdefault(key, set()).add(a)
    print(f"OTA names from {Path(reservation_report).name}")

    def ota_of(guest):
        otas = found.get(" ".join(guest.split()).lower(), set())
        return next(iter(otas)) if len(otas) == 1 else None  # 2+ OTAs: can't tell which
    return ota_of


MANUAL_FILE = "ota_missing_details.json"  # answers typed in the "missing OTA" form


def load_manual():
    """{hotel: {Mews guest name: {"ota": ..., "name": ...}}} saved from the form."""
    import json
    f = HERE / MANUAL_FILE
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except ValueError:
        return {}


def save_manual(enterprise, answers):
    import json
    data = load_manual()
    data.setdefault(enterprise, {}).update(answers)
    (HERE / MANUAL_FILE).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


def customer_name(guest, ota_of, manual):
    """'OTA:Name' for a Mews guest: Reservation report first, then the form's answers.
    Returns (name, found) - found is False when there's still no OTA."""
    m = manual.get(guest, {})
    ota = ota_of(guest) or m.get("ota")
    name = m.get("name") or guest
    return (f"{ota}:{name}" if ota else name), bool(ota)


def guests_without_ota(report_path):
    """Before generating: guests (with their cash) who'd get no OTA:parent prefix."""
    import contextlib
    import io
    enterprise, cash = read_cash(report_path, quiet=True)
    with contextlib.redirect_stdout(io.StringIO()):
        ota_of = ota_lookup(find_reservation_report(report_path, enterprise))
    manual = load_manual().get(enterprise, {})
    missing = {}
    for guest, value in zip(cash["Customer"].astype(str).str.strip(), cash["Value"]):
        if not customer_name(guest, ota_of, manual)[1]:
            missing[guest] = missing.get(guest, 0) + int(value)
    return missing


RECON_ACCOUNTS = ["XXXXXXX Recon Dis. / Round. Difference", "XXXXXXX Reconciliation Discrepancies"]


def discrepancy(net, counted, account):
    """Before generating: the discrepancy line a payout would create, or None."""
    if counted is None or counted == net:
        return None
    diff = counted - net
    return {"amount": abs(diff), "kind": "over" if diff > 0 else "short",
            "side": "Credit" if diff > 0 else "Debit", "account": account}


def known_otas(report_path):
    """OTA choices for the form: what the Reservation report uses, plus Direct and Expedia."""
    enterprise, _ = read_cash(report_path, quiet=True)
    res = find_reservation_report(report_path, enterprise)
    found = set()
    if res:
        agency = pd.read_excel(res, sheet_name="Reservations")["Travel agency"].dropna().astype(str).str.strip()
        found = {OTA_NAMES.get(a, a) for a in agency if a}
    return [DIRECT] + sorted(found | {"Booking.com", "Agoda", "Expedia"})


def upload_name(journal_date, enterprise, doc_no):
    """e.g. 2026-09-01_OriIkebukuro_Accommodation_PettyCash_134.xlsx (journal entry date)."""
    hotel = (str(master_row(enterprise).get("Evidence Naming") or "").strip()      # e.g. WFGojo
             or "".join(w.capitalize() for w in re.split(r"[^A-Za-z0-9]+", enterprise) if w))
    return f"{journal_date:%Y-%m-%d}_{hotel}{UPLOAD_TAG}_{doc_no}.xlsx"


def read_cash(report_path, quiet=False):
    """Cash rows from the External payments tab, oldest first."""
    params = pd.read_excel(report_path, sheet_name="Parameters", header=None)
    enterprise = str(params.loc[params[0] == "Enterprise", 1].iloc[0]).strip()
    df = pd.read_excel(report_path, sheet_name="External payments")
    cash = df[df["Type"] == "Cash"].sort_values("Created")
    # Test bookings (e.g. "Test631 Test631") aren't real money; leave them out.
    is_test = cash["Customer"].astype(str).str.contains(r"\btest\d*\b", case=False, regex=True)
    for _, p in ([] if quiet else cash[is_test].iterrows()):
        print(f"Skipped test entry: {p['Customer']} {int(p['Value']):,}")
    return enterprise, cash[~is_test]


def parse_extra(text):
    """'Name=29000, Other Name=-500' -> [(name, amount), ...]"""
    extra = []
    for part in filter(None, (s.strip() for s in (text or "").split(","))):
        name, _, amount = part.rpartition("=")
        if not name.strip():
            raise ValueError(f"'{part}' has no guest name")
        extra.append((name.strip(), int(amount.replace(" ", "").replace("¥", ""))))
    return extra


def copy_to_documentation(out, tab, journal_date):
    """Copy the upload file to Documentation \\ YYYY.MM \\ Petty Cash \\ <hotel folder>."""
    import shutil
    root = Path(DOC_ROOT)
    if not root.exists():
        return f"NOT copied to Documentation ({DOC_ROOT} not found)"
    folder = DOC_FOLDERS.get(tab.strip())
    if not folder:
        return f"NOT copied to Documentation (no folder set for tab {tab.strip()} in DOC_FOLDERS)"
    base = root / f"{journal_date:%Y.%m}" / "Petty Cash"
    bare = lambda n: re.sub(r"^\d+\.\s*", "", n).strip().lower()  # noqa: E731
    existing = [d for d in base.glob("*") if d.is_dir() and bare(d.name) == bare(folder)] if base.exists() else []
    dest = existing[0] if existing else base / folder
    dest.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out, dest / Path(out).name)
    return f"Copied to Documentation: {journal_date:%Y.%m}\\Petty Cash\\{dest.name}"


def build(report_path, tab, doc_no, journal_date, template=TEMPLATE, out_dir=None,
          counted=None, extra=None, remove_reports=True, recon_acct=None):
    out_dir = out_dir or EVIDENCE
    enterprise, cash = read_cash(report_path)
    # Cash that Mews recorded under another method (e.g. as card): add by hand.
    if extra:
        manual = pd.DataFrame([{"Customer": n, "Value": v, "Accounting category": "MANUAL"}
                               for n, v in extra])
        cash = pd.concat([cash, manual], ignore_index=True)
    if cash.empty:
        raise SystemExit("No cash payments found in this report.")
    net = int(cash["Value"].sum())

    tab = resolve_tab(tab, template)
    wb = openpyxl.load_workbook(template)
    recon = (recon_acct or recon_account(wb, wb[tab])) if counted is not None and counted != net else None
    if counted is not None and counted != net and not recon:
        recon = ask("Reconciliation discrepancy account for this property")
    # Upload file holds just this hotel's tab.
    for name in wb.sheetnames:
        if name != tab:
            del wb[name]
    ws = wb[tab]
    ws.sheet_view.tabSelected = True

    # Row 2 holds the fixed property details; reuse them for every line.
    cls, loc = ws["I2"].value, ws["J2"].value
    desc, memo = ws["K2"].value, ws["Q2"].value
    line_style = [copy(ws.cell(3, c)._style) for c in range(1, 24)]

    for r in range(3, ws.max_row + 1):
        for c in range(1, 24):
            ws.cell(r, c).value = None

    # Plain values, not formulas: Business Importer reads a formula cell as blank
    # unless Excel has recalculated it, which splits every line into its own entry.
    total = counted if counted is not None else net
    ws["A2"], ws["D2"], ws["R2"], ws["C2"] = doc_no, journal_date, journal_date, total

    ota_of = ota_lookup(find_reservation_report(report_path, enterprise))
    manual_answers = load_manual().get(enterprise, {})

    flagged = []
    for r, (_, p) in enumerate(cash.iterrows(), start=3):
        amount = int(p["Value"])
        guest = str(p["Customer"]).strip()
        name, ota = customer_name(guest, ota_of, manual_answers)
        values = {
            1: doc_no, 2: "Accounts Receivable (A/R)", 3: abs(amount), 4: journal_date,
            5: name, 6: "JPY", 9: cls, 10: loc, 11: desc,
            12: "Credit" if amount >= 0 else "Debit",  # assumption: refunds are negative
            17: memo, 18: journal_date,
        }
        for c in range(1, 24):
            cell = ws.cell(r, c)
            cell._style = copy(line_style[c - 1])
            if c in values:
                cell.value = values[c]
        manual = p["Accounting category"] == "MANUAL"
        if pd.isna(p["Accounting category"]) or amount < 0 or manual or not ota:
            ws.cell(r, 3).fill = ws.cell(r, 5).fill = YELLOW
            notes = [n for n, on in ((" (added by hand)", manual),
                                     (" (OTA unknown - add it by hand)", not ota)) if on]
            flagged.append(f"  row {r}: {guest} {amount:,}{''.join(notes)}")

    # Counted cash differs from Mews: book the gap to the discrepancy account.
    # Short (counted < Mews) is a Debit, over is a Credit.
    diff = (counted - net) if counted is not None else 0
    if diff:
        r = len(cash) + 3
        values = {
            1: doc_no, 2: recon, 3: abs(diff), 4: journal_date, 6: "JPY", 9: cls, 10: loc,
            11: f"Reconciliation discrepancy {desc}", 12: "Credit" if diff > 0 else "Debit",
            17: f"Reconciliation discrepancy {memo}", 18: journal_date,
        }
        for c in range(1, 24):
            cell = ws.cell(r, c)
            cell._style = copy(line_style[c - 1])
            if c in values:
                cell.value = values[c]
        ws.cell(r, 2).fill = ws.cell(r, 3).fill = YELLOW
        flagged.append(f"  row {r}: reconciliation discrepancy {'over' if diff > 0 else 'short'} "
                       f"{abs(diff):,} ({values[12]}) to {recon}")

    # Drop the template's leftover empty rows so nothing blank gets imported.
    last = len(cash) + 2 + (1 if diff else 0)
    if ws.max_row > last:
        ws.delete_rows(last + 1, ws.max_row - last)
    for extra_row in [n for n in ws.row_dimensions if n > last]:
        del ws.row_dimensions[extra_row]  # formatted-but-empty rows still count as "used"

    out = Path(out_dir) / upload_name(journal_date, enterprise, doc_no)
    wb.save(out)
    if not excel_resave(out):
        use_shared_strings(out)
    removed = remove_old_uploads(out)

    print(f"\n{enterprise}: {len(cash)} cash lines, net JPY {net:,}")
    if diff:
        print(f"Counted JPY {counted:,}: added {'over' if diff > 0 else 'short'} {abs(diff):,} "
              f"({values[12]}) to {recon} on row {r}. Total is now JPY {counted:,}.")
    print(f"Saved {out.name}")
    try:
        print(copy_to_documentation(out, tab, journal_date))
    except Exception as e:
        print(f"NOT copied to Documentation ({e})")
    for name in removed:
        print(f"Moved old upload file to the Recycle Bin: {name}")
    if remove_reports:
        for name in remove_used_reports(report_path, enterprise):
            print(f"Moved used report to the Recycle Bin: {name}")
    if flagged:
        print("Check these highlighted lines (no category in Mews, a refund, added by hand, "
              "OTA unknown, or discrepancy):")
        print("\n".join(flagged))
    if counted is None:
        print("Compare the net total with the counted cash before uploading.")
    return out


def main():
    if len(sys.argv) == 1:  # double-click / .bat: open the window
        run_window()
        return
    ap = argparse.ArgumentParser()
    ap.add_argument("--report")
    ap.add_argument("--tab")
    ap.add_argument("--doc")
    ap.add_argument("--date", help="QuickBooks journal date, YYYY-MM-DD")
    ap.add_argument("--counted", help="counted cash in JPY; any gap goes to the discrepancy account")
    ap.add_argument("--extra", help='cash missing from the report, e.g. "Jane Doe=29000"')
    ap.add_argument("--template", default=TEMPLATE)
    a = ap.parse_args()

    report = clean(a.report) if a.report else pick_report()
    enterprise, _ = read_cash(report, quiet=True)
    found = find_tab(enterprise, a.template)
    if a.tab or found:
        tab = a.tab or found
        print(f"{enterprise} -> tab {tab.strip()}")
    else:
        tab = ask(f"No tab found for {enterprise}. Type the template tab name")
    if not sys.stdin.isatty() and not (a.doc and a.date):
        print("Tip: when running without typing answers, add the Doc No and date, e.g.\n"
              "  py petty_cash_bi.py --doc 132 --date 9/1")
    doc = retry(lambda: int(clean(a.doc or ask("Doc No (next QuickBooks journal entry)")).replace(",", "")),
                "Doc No must be a number.")
    date = retry(lambda: parse_date(a.date or ask("QuickBooks journal date (e.g. 9/1 or 2026-09-01)")),
                 "Couldn't read that date. Try 9/1 or 2026-09-01.")
    extra = parse_extra(a.extra)  # only via --extra; not asked
    counted = retry(lambda: parse_amount(a.counted or ask(
                        "Payout in JPY - any difference goes to discrepancy (Enter if it matches)")),
                    "Payout must be a number, e.g. 63888.")
    no_ota = guests_without_ota(report)
    if no_ota:
        print(f"\nWARNING: {len(no_ota)} guest(s) have no OTA (no parent customer):")
        for g, v in no_ota.items():
            print(f"  {g}  {v:,}")
    build(report, tab, doc, date, template=a.template, counted=counted, extra=extra)


def clean(text):
    return str(text or "").replace("﻿", "").strip().strip('"').strip()


def parse_date(text):
    """Accepts 2026-09-01, 9/1, 09/01, 9/1/2026 (month first); no year means this year."""
    text = clean(text)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            pass
    month, day = text.split("/")
    return dt.datetime(dt.date.today().year, int(month), int(day))


def parse_amount(text):
    text = clean(text).replace(",", "").replace("¥", "")
    return int(text) if text else None


def retry(read, message):
    """Ask again on a bad answer instead of crashing (only when typing interactively)."""
    while True:
        try:
            return read()
        except (ValueError, TypeError):
            if not sys.stdin.isatty():
                raise
            print(message)


def pick_report():
    """Offer the payment reports in this folder by number, newest first."""
    reports = payment_reports()
    if not reports:
        return clean(ask("Mews payment report file (drag it here)"))
    if len(reports) == 1:
        print(f"Report: {reports[0].name}")
        return str(reports[0])
    print("Payment reports in this folder:")
    for i, f in enumerate(reports, 1):
        print(f"  {i}. {f.name}")
    while True:
        answer = clean(ask("Pick a number, or drag a report here", "1"))
        if answer.isdigit() and 1 <= int(answer) <= len(reports):
            return str(reports[int(answer) - 1])
        if Path(answer).exists():
            return answer
        print("Didn't find that. Type one of the numbers above.")


# ---------------------------------------------------------------------------
# Window (same look as the bank recon OTA JE tool): pick the date from a calendar
# ---------------------------------------------------------------------------
def payment_reports(folder=None):
    files = sorted(Path(folder or EVIDENCE).glob("Payment*report*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    return [f for f in files if not f.name.startswith("~$")]


def next_doc_number(folder=None):
    """Doc number of the last upload file in the folder + 1, or ''."""
    nums = [int(m.group(1)) for f in Path(folder or EVIDENCE).glob(f"*{UPLOAD_TAG}_*.xlsx")
            if (m := re.search(r"_(\d+)\.xlsx$", f.name))]
    return str(max(nums) + 1) if nums else ""


def run_window():
    import calendar
    import contextlib
    import io
    import os
    import tkinter as tk
    from tkinter import ttk

    green, light, text, red = "#C6E0B4", "#EEF6E9", "#375623", "#9C0006"
    root = tk.Tk()
    root.title("Petty Cash - Business Importer upload")
    root.configure(bg="white")
    root.resizable(False, False)

    # Copy/paste that works in every text box:
    # - Ctrl+C/V/X/A by the physical key, so they still work with Japanese input (IME) on
    #   or Caps Lock, where Tk can miss the normal Ctrl+V.
    # - Clicking a box always takes the keyboard back (after switching to Mews and back,
    #   Windows can leave the keyboard on the locked main window behind a pop-up).
    # - Right-click menu: Cut / Copy / Paste / Select all.
    VK = {86: "<<Paste>>", 67: "<<Copy>>", 88: "<<Cut>>", 65: "<<SelectAll>>"}

    def ctrl_key(e):
        if e.keycode in VK:
            e.widget.event_generate(VK[e.keycode])
            return "break"

    def take_focus(e):
        e.widget.winfo_toplevel().focus_force()
        e.widget.focus_set()

    edit_menu = tk.Menu(root, tearoff=0)

    def context_menu(e):
        w = e.widget
        take_focus(e)
        edit_menu.delete(0, "end")
        for label, ev in (("Cut", "<<Cut>>"), ("Copy", "<<Copy>>"), ("Paste", "<<Paste>>"),
                          ("Select all", "<<SelectAll>>")):
            edit_menu.add_command(label=label, command=lambda ev=ev: w.event_generate(ev))
        edit_menu.tk_popup(e.x_root, e.y_root)

    for cls in ("Entry", "TEntry", "TCombobox"):
        root.bind_class(cls, "<Control-KeyPress>", ctrl_key, add="+")
        root.bind_class(cls, "<Button-1>", take_focus, add="+")
        root.bind_class(cls, "<Button-3>", context_menu, add="+")

    # Coming back to the app (taskbar / Alt+Tab) while a pop-up is open: give it the keyboard.
    open_form = {"w": None}

    def back_to_form(_):
        w = open_form["w"]
        if w is not None and w.winfo_exists() and root.focus_get() in (None, root):
            w.lift()
            w.focus_force()
    root.bind("<FocusIn>", back_to_form, add="+")

    def pick_date(var, anchor):
        """Small calendar popup; clicking a day fills the date."""
        try:
            cur = dt.datetime.strptime(var.get(), "%Y/%m/%d")
        except ValueError:
            cur = dt.datetime.today()
        state = {"y": cur.year, "m": cur.month}
        top = tk.Toplevel(root, bg="white")
        top.title("Pick a date")
        top.transient(root)
        top.grab_set()
        top.geometry(f"+{anchor.winfo_rootx()}+{anchor.winfo_rooty() + anchor.winfo_height()}")
        head = tk.Frame(top, bg=green)
        head.pack(fill="x")
        title = tk.Label(head, bg=green, fg=text, font=("Segoe UI", 10, "bold"), width=16)
        grid = tk.Frame(top, bg="white", padx=6, pady=6)
        grid.pack()

        def draw():
            for w in grid.winfo_children():
                w.destroy()
            title.config(text=f"{calendar.month_name[state['m']]} {state['y']}")
            for i, d in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
                tk.Label(grid, text=d, bg="white", fg="#7F7F7F", font=("Segoe UI", 8)).grid(row=0, column=i)
            for r, week in enumerate(calendar.monthcalendar(state["y"], state["m"]), 1):
                for c, d in enumerate(week):
                    if not d:
                        continue
                    sel = (d == cur.day and state["m"] == cur.month and state["y"] == cur.year)
                    tk.Button(grid, text=str(d), width=3, relief="flat",
                              bg=green if sel else light, fg=text, activebackground=green,
                              command=lambda d=d: choose(d)).grid(row=r, column=c, padx=1, pady=1)

        def move(step):
            m = state["m"] + step
            state["y"] += (m - 1) // 12
            state["m"] = (m - 1) % 12 + 1
            draw()

        def choose(d):
            var.set(f"{state['y']}/{state['m']:02d}/{d:02d}")
            top.destroy()

        tk.Button(head, text="◀", relief="flat", bg=green, command=lambda: move(-1)).pack(side="left")
        title.pack(side="left")
        tk.Button(head, text="▶", relief="flat", bg=green, command=lambda: move(1)).pack(side="left")
        draw()

    def review(report, missing, disc, fill_only=False):
        """Form shown before generating: the reconciliation discrepancy (if any) and
        guests with no OTA. Returns ('save'|'skip'|'cancel', discrepancy account).
        fill_only: opened from the 'Fill in' button - just Save / Close, no generating."""
        enterprise, _ = read_cash(report, quiet=True)
        saved = load_manual().get(enterprise, {})
        choices = known_otas(report)
        pink = "#FFC7CE"
        top = tk.Toplevel(root, bg="white")
        top.title("Check before generating")
        open_form["w"] = top
        top.transient(root)
        top.grab_set()
        avar = tk.StringVar(value=disc["account"] if disc else "")
        if disc:
            tk.Label(top, text="Reconciliation discrepancy - payout differs from Mews",
                     bg="white", fg=text, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))
            dg = tk.Frame(top, bg="white")
            dg.pack(anchor="w", padx=10)
            for c, h in enumerate(["Line", "Amount", "Over / short", "Debit / Credit", "Account"]):
                tk.Label(dg, text=h, bg=green, fg=text, font=("Segoe UI", 9, "bold"),
                         padx=6).grid(row=0, column=c, sticky="ew")
            for c, v in enumerate(["Reconciliation discrepancy", f"¥{disc['amount']:,}",
                                   disc["kind"], disc["side"]]):
                tk.Label(dg, text=v, bg=pink, fg=red, padx=6).grid(row=1, column=c, sticky="ew")
            ttk.Combobox(dg, textvariable=avar, values=RECON_ACCOUNTS, width=40).grid(row=1, column=4, padx=3)
        if missing:
            tk.Label(top, text="Guests with no OTA in the Reservation report - pick the OTA "
                     "(and correct the name if needed). What you enter is saved and reused next time.",
                     bg="white", fg=text, font=("Segoe UI", 10, "bold"), wraplength=640,
                     justify="left").pack(anchor="w", padx=10, pady=(12, 4))
        outer = tk.Frame(top, bg="white")
        outer.pack(fill="both", expand=True, padx=10)
        canvas = tk.Canvas(outer, bg="white", highlightthickness=0,
                           height=min(40 + 31 * len(missing), 480) if missing else 1)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        grid = tk.Frame(canvas, bg="white")
        grid.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=grid, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        if len(missing) > 14:
            sb.pack(side="right", fill="y")
        for c, h in enumerate(["Customer (Mews)", "Cash", "OTA", "Name for QuickBooks", "Preview"]
                              if missing else []):
            tk.Label(grid, text=h, bg=green, fg=text, font=("Segoe UI", 9, "bold"),
                     padx=6).grid(row=0, column=c, sticky="ew")
        rows_ = []
        for i, (guest, amount) in enumerate(missing.items(), 1):
            bg = light if i % 2 == 0 else "white"
            prev = saved.get(guest, {})
            tk.Label(grid, text=guest, bg=bg, anchor="w", padx=6).grid(row=i, column=0, sticky="ew")
            tk.Label(grid, text=f"¥{amount:,}", bg=bg, anchor="e", padx=6).grid(row=i, column=1, sticky="ew")
            ovar = tk.StringVar(value=prev.get("ota", ""))
            nvar_ = tk.StringVar(value=prev.get("name", guest))
            ob = ttk.Combobox(grid, textvariable=ovar, values=choices, width=16)
            ob.grid(row=i, column=2, padx=3, pady=2)
            tk.Entry(grid, textvariable=nvar_, width=28).grid(row=i, column=3, padx=3, pady=2)
            pv = tk.Label(grid, text="", bg=bg, anchor="w", padx=6, width=34)
            pv.grid(row=i, column=4, sticky="ew")

            def update(*_, ovar=ovar, nvar_=nvar_, pv=pv):
                o, n = ovar.get().strip(), nvar_.get().strip()
                pv.config(text=f"{o}:{n}" if o else "(no OTA - stays yellow)",
                          fg=text if o else "#9C0006", bg=bg if o else pink)
            ovar.trace_add("write", update)
            nvar_.trace_add("write", update)
            update()
            rows_.append((guest, ovar, nvar_))
        choice = {"v": "cancel"}

        def save():
            answers = {}
            for guest, ovar, nvar_ in rows_:
                o, n = ovar.get().strip(), nvar_.get().strip()
                if o:
                    answers[guest] = {"ota": o, **({"name": n} if n and n != guest else {})}
            if answers:
                save_manual(enterprise, answers)
            choice["v"] = "save"
            top.destroy()

        def skip():
            choice["v"] = "skip"
            top.destroy()

        bar_ = tk.Frame(top, bg="white")
        bar_.pack(anchor="e", padx=10, pady=10)
        tk.Button(bar_, text="Save" if fill_only else "Save & generate" if missing else "Confirm & generate",
                  bg=green, fg=text, relief="flat", padx=12, pady=4, font=("Segoe UI", 10, "bold"),
                  command=save).pack(side="left", padx=4)
        if missing and not fill_only:
            tk.Button(bar_, text="Skip OTA - generate with yellow lines", bg=light, relief="flat",
                      padx=10, pady=4, command=skip).pack(side="left", padx=4)
        tk.Button(bar_, text="Cancel", bg=light, relief="flat", padx=10, pady=4,
                  command=top.destroy).pack(side="left", padx=4)
        top.wait_window()
        open_form["w"] = None
        return choice["v"], avar.get().strip()

    tk.Label(root, text="Petty cash upload", bg="white", fg=text,
             font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 4))

    def label(row, txt):
        tk.Label(root, text=txt, bg=green, fg=text, font=("Segoe UI", 9, "bold"), padx=6,
                 anchor="w", width=16).grid(row=row, column=0, sticky="ew", padx=(10, 4), pady=2)

    label(1, "Payment report")
    reports = payment_reports()
    rvar = tk.StringVar(value=reports[0].name if reports else "")
    box = ttk.Combobox(root, textvariable=rvar, values=[f.name for f in reports], width=58, state="readonly")
    box.grid(row=1, column=1, columnspan=3, sticky="w", pady=2)

    info = tk.Label(root, text="", bg="white", fg=text, anchor="w", justify="left")
    info.grid(row=2, column=1, columnspan=3, sticky="w", pady=(0, 6))
    hotel = {"tab": None}

    prompted = set()  # reports whose form already popped up by itself

    def fill_in():
        if hotel.get("no_ota"):
            prompted.add(rvar.get())
            review(EVIDENCE / rvar.get(), hotel["no_ota"], None, fill_only=True)
            show_report()  # re-check: filled-in guests drop off the list

    fill_btn = tk.Button(root, text="✎ Fill in missing details", bg="#FFC7CE", fg=red, relief="flat",
                         font=("Segoe UI", 9, "bold"), padx=8, pady=3, command=fill_in)

    def show_report(*_):
        fill_btn.grid_remove()
        hotel.clear()
        hotel["tab"], hotel["no_ota"] = None, {}
        if not rvar.get():
            info.config(text="No Payment report in this folder. Put the Mews reports here.", fg=red)
            return
        try:
            path = EVIDENCE / rvar.get()
            enterprise, cash = read_cash(path, quiet=True)
            hotel["tab"] = find_tab(enterprise)
            hotel["net"] = int(cash["Value"].sum())
            if hotel["tab"]:
                twb = openpyxl.load_workbook(TEMPLATE)
                hotel["recon"] = recon_account(twb, twb[hotel["tab"]])
            res = find_reservation_report(path, enterprise)
            lines = [f"{enterprise}  ->  tab {hotel['tab'].strip() if hotel['tab'] else 'NOT FOUND in template'}",
                     f"{len(cash)} cash lines, net ¥{int(cash['Value'].sum()):,}",
                     f"Reservation report: {res.name if res else 'MISSING - OTA names will be blank'}"]
            hotel["no_ota"] = guests_without_ota(path) if res else {}
            if hotel["no_ota"]:
                lines.append(f"{len(hotel['no_ota'])} guest(s) with NO OTA (no parent customer) - see below")
            bad = not hotel["tab"] or not res or hotel["no_ota"]
            info.config(text="\n".join(lines), fg=red if bad else text)
            if hotel["no_ota"]:
                fill_btn.grid(row=2, column=0, sticky="nw", padx=(10, 4), pady=(2, 6))
                if rvar.get() not in prompted:  # pop the form up straight away, once per report
                    root.after(300, fill_in)
                write("Check before generating - these guests have no OTA, so they'd be uploaded "
                      "without a parent customer:\n\n"
                      + "\n".join(f"  {g}   ¥{v:,}" for g, v in hotel["no_ota"].items())
                      + "\n\nNot in the Reservation report, or booked under 2+ OTAs. Click "
                        "'Fill in missing details' (left) to enter the OTA and fix the name.", red)
            else:
                fill_btn.grid_remove()
                write("")
        except Exception as e:
            info.config(text=f"Can't read this report: {e}", fg=red)
        show_discrepancy()

    box.bind("<<ComboboxSelected>>", show_report)

    label(3, "Journal date")
    dvar = tk.StringVar(value=dt.date.today().strftime("%Y/%m/%d"))
    dbox = tk.Frame(root, bg="white")
    dbox.grid(row=3, column=1, columnspan=3, sticky="w", pady=2)
    ttk.Entry(dbox, textvariable=dvar, width=12).pack(side="left")
    cal = tk.Button(dbox, text="📅", relief="flat", bg="white")
    cal.config(command=lambda: pick_date(dvar, cal))
    cal.pack(side="left", padx=4)

    label(4, "Doc number")
    nvar = tk.StringVar(value=next_doc_number())
    ttk.Entry(root, textvariable=nvar, width=12).grid(row=4, column=1, sticky="w", pady=2)

    label(5, "Payout (¥)")
    pvar = tk.StringVar(value="")
    pbox = tk.Frame(root, bg="white")
    pbox.grid(row=5, column=1, columnspan=3, sticky="w", pady=2)
    ttk.Entry(pbox, textvariable=pvar, width=12).pack(side="left")
    disc_lbl = tk.Label(pbox, text="", bg="white", fg="#7F7F7F")
    disc_lbl.pack(side="left", padx=6)

    def current_discrepancy():
        try:
            counted = parse_amount(pvar.get())
        except ValueError:
            return "bad"
        if "net" not in hotel:
            return None
        return discrepancy(hotel["net"], counted, hotel.get("recon") or RECON_ACCOUNTS[0])

    def show_discrepancy(*_):
        d = current_discrepancy()
        if d == "bad":
            disc_lbl.config(text="payout must be a number", fg=red)
        elif d:
            disc_lbl.config(text=f"⚠ Discrepancy: {d['kind']} ¥{d['amount']:,} -> {d['side']} to "
                                 f"{d['account']} (you'll confirm it before generating)", fg=red)
        elif pvar.get().strip():
            disc_lbl.config(text="✔ matches Mews - no discrepancy", fg=text)
        else:
            disc_lbl.config(text="counted cash - any difference goes to discrepancy; leave blank if it matches",
                            fg="#7F7F7F")

    pvar.trace_add("write", show_discrepancy)

    log = tk.Text(root, width=96, height=16, relief="flat", bg=light, fg="#222222", font=("Consolas", 9))
    log.grid(row=6, column=0, columnspan=4, padx=10, pady=(8, 4))
    log.config(state="disabled")

    def write(msg, colour=None):
        log.config(state="normal")
        log.delete("1.0", "end")
        log.insert("end", msg)
        log.config(state="disabled", fg=colour or "#222222")

    def generate():
        try:
            if not rvar.get():
                raise ValueError("pick a Payment report")
            if not hotel["tab"]:
                raise ValueError("this hotel has no tab in the template - add it first")
            counted = parse_amount(pvar.get())
        except ValueError as e:
            write(f"✖ {e}", red)
            return
        # Missing details / discrepancy first, so the form always shows.
        disc = current_discrepancy()
        recon_acct = None
        if hotel.get("no_ota") or disc:
            choice, recon_acct = review(EVIDENCE / rvar.get(), hotel.get("no_ota") or {}, disc)
            if choice == "cancel":
                write("Cancelled - nothing generated.", red)
                return
        try:
            date = dt.datetime.strptime(dvar.get().strip(), "%Y/%m/%d")
            doc = nvar.get().strip().replace(",", "")
            if not doc.isdigit():
                raise ValueError("enter the doc number, then click Generate again "
                                 "(what you filled in is saved)")
        except ValueError as e:
            write(f"✖ {e}", red)
            return
        write("Working... (Excel saves the file, this takes a few seconds)")
        root.config(cursor="watch")
        root.update()
        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                build(EVIDENCE / rvar.get(), hotel["tab"], int(doc), date, counted=counted,
                      recon_acct=recon_acct or None)
            out = buf.getvalue().strip()
            write("✔ Done\n\n" + out, red if "NOT copied" in out else text)
            nvar.set(str(int(doc) + 1))
            refresh_list()  # the used report is gone now
        except BaseException as e:  # SystemExit from build is a message too
            write(buf.getvalue() + f"\n✖ {e}", red)
        finally:
            root.config(cursor="")

    def refresh_list():
        nonlocal reports
        reports = payment_reports()
        box.config(values=[f.name for f in reports])
        if rvar.get() not in [f.name for f in reports]:
            rvar.set(reports[0].name if reports else "")
        show_report()

    refresh = refresh_list

    bar = tk.Frame(root, bg="white")
    bar.grid(row=7, column=0, columnspan=4, sticky="e", padx=10, pady=10)
    tk.Button(bar, text="Generate", bg=green, fg=text, relief="flat", padx=14, pady=4,
              font=("Segoe UI", 10, "bold"), command=generate).pack(side="left", padx=4)
    tk.Button(bar, text="Refresh reports", relief="flat", bg=light, padx=10, pady=4,
              command=refresh).pack(side="left", padx=4)
    tk.Button(bar, text="Open folder", relief="flat", bg=light, padx=10, pady=4,
              command=lambda: os.startfile(EVIDENCE)).pack(side="left", padx=4)
    tk.Button(bar, text="Close", relief="flat", bg=light, padx=10, pady=4,
              command=root.destroy).pack(side="left", padx=4)

    show_report()
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except (SystemExit, KeyboardInterrupt):
        raise
    except Exception as e:  # show a plain message instead of a Python traceback
        print(f"\nSomething went wrong: {e or type(e).__name__}")
        print("Nothing was saved. Send a screenshot of this window if it keeps happening.")
        sys.exit(1)

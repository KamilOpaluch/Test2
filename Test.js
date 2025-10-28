# access_table_sizes.py
import os
import sys
import csv
from datetime import datetime

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    raise SystemExit("Requires pywin32. Install with: pip install pywin32")

# ---- Estimation tunables ----
ROW_OVERHEAD_BYTES = 24
INDEX_OVERHEAD_MULT = 1.10
DEFAULT_AVG_MEMO_CHARS = 200
DEFAULT_TEXT_FILL = 0.5

# ---- DAO constants ----
DB_BOOLEAN = 1
DB_BYTE = 2
DB_INTEGER = 3
DB_LONG = 4
DB_CURRENCY = 5
DB_SINGLE = 6
DB_DOUBLE = 7
DB_DATE = 8
DB_BINARY = 9
DB_TEXT = 10
DB_LONG_BINARY = 11
DB_MEMO = 12
DB_GUID = 15
DB_DECIMAL = 20
DB_ATTACHMENT = 101

DB_OPEN_TABLE = 1
DB_OPEN_SNAPSHOT = 4

# ----------------- Access attach helpers -----------------

def enumerate_access_instances():
    """Return a list of running Access.Application COM objects by enumerating the ROT."""
    apps = []
    ctx = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    enum = rot.EnumRunning()
    while True:
        fetched = enum.Next(1)
        if not fetched:
            break
        mk = fetched[0]
        try:
            name = mk.GetDisplayName(ctx, None)
        except pythoncom.com_error:
            continue
        if "Access.Application" in name or name.lower().endswith("!{73fddc80-aea9-101a-98a7-00aa00374959}"):
            try:
                obj = rot.GetObject(mk)   # IUnknown
                app = win32.Dispatch(obj) # Wrap for late-binding
                apps.append(app)
            except Exception:
                pass
    return apps

def get_active_db_path_from_app(app):
    """Try both Access object model paths; return (db_handle, full_path, base_name) or (None, '', '')."""
    # Try CurrentProject.FullName
    try:
        full = str(app.CurrentProject.FullName)
        if full:
            try:
                db = app.CurrentDb()
                _ = db.TableDefs.Count
                return db, full, os.path.basename(full)
            except Exception:
                pass
    except Exception:
        pass
    # Try via CurrentDb().Name
    try:
        db = app.CurrentDb()
        full = str(db.Name)
        if full:
            _ = db.TableDefs.Count
            return db, full, os.path.basename(full)
    except Exception:
        pass
    return None, "", ""

def open_db_via_engine(app, full_path):
    """Open a DAO.Database via the instance's DBEngine."""
    ws0 = app.DBEngine.Workspaces(0)
    db = ws0.OpenDatabase(full_path)
    _ = db.TableDefs.Count  # sanity
    return db

def pick_database(name_fragment="", explicit_path=None):
    """
    Find a running Access instance with an open DB that matches name_fragment.
    If explicit_path is given, open that path via any found instance's DBEngine.
    """
    apps = enumerate_access_instances()
    if not apps:
        raise RuntimeError("No running Access.Application instances found. Open your .accdb and try again.")

    # If an explicit path is provided, try to open it via the first instance's DBEngine.
    if explicit_path:
        explicit_path = os.path.abspath(explicit_path)
        for app in apps:
            try:
                db = open_db_via_engine(app, explicit_path)
                return app, db, explicit_path, os.path.basename(explicit_path)
            except Exception:
                continue
        raise RuntimeError(f"Could not open database via DAO at: {explicit_path}")

    # Otherwise, search each instance for an already-open DB
    candidates = []
    for app in apps:
        db, full, base = get_active_db_path_from_app(app)
        if db and full:
            candidates.append((app, db, full, base))

    if not candidates:
        raise RuntimeError("Found Access instances, but none reported an open database "
                           "(window without file, or insufficient permissions).")

    # Prefer a candidate whose filename contains the fragment
    if name_fragment:
        frag = name_fragment.lower()
        for app, db, full, base in candidates:
            if frag in base.lower():
                return app, db, full, base

    # Fallback: return the first open DB and warn
    app, db, full, base = candidates[0]
    if name_fragment:
        print(f"[warn] No open database matching '{name_fragment}'. Using '{base}'.")
    return app, db, full, base

# ----------------- DAO helpers -----------------

def iter_collection(coll):
    cnt = int(coll.Count)
    # try 0-based
    try:
        for i in range(cnt):
            yield coll.Item(i)
        return
    except Exception:
        pass
    # fallback 1-based
    for i in range(1, cnt + 1):
        yield coll.Item(i)

def is_user_local_table(tdef):
    nm = str(tdef.Name)
    if nm.startswith("MSys"):
        return False
    try:
        if len(str(tdef.Connect)) > 0:
            return False  # linked
    except Exception:
        pass
    return True

def field_size_bytes(fld, avg_memo_chars=DEFAULT_AVG_MEMO_CHARS, text_fill=DEFAULT_TEXT_FILL):
    ftype = int(fld.Type)
    try:
        fsize = int(fld.Size)
    except Exception:
        fsize = 0

    if ftype == DB_BOOLEAN:  return 1
    if ftype == DB_BYTE:     return 1
    if ftype == DB_INTEGER:  return 2
    if ftype == DB_LONG:     return 4
    if ftype == DB_SINGLE:   return 4
    if ftype == DB_DOUBLE:   return 8
    if ftype == DB_CURRENCY: return 8
    if ftype == DB_DECIMAL:  return 12
    if ftype == DB_DATE:     return 8
    if ftype == DB_GUID:     return 16
    if ftype == DB_BINARY:   return max(fsize, 0)
    if ftype == DB_TEXT:     return int(round((fsize * 2) * text_fill))
    if ftype == DB_MEMO:     return int(avg_memo_chars * 2)
    if ftype in (DB_LONG_BINARY, DB_ATTACHMENT):
        return 0
    return 0

def estimated_row_bytes(tdef, avg_memo_chars, text_fill):
    total = 0
    for fld in iter_collection(tdef.Fields):
        total += field_size_bytes(fld, avg_memo_chars, text_fill)
    return total

def fast_row_count(db, table_name):
    # Try dbOpenTable
    try:
        rs = db.OpenRecordset(table_name, DB_OPEN_TABLE)
        rc = int(rs.RecordCount)
        rs.Close()
        return rc
    except Exception:
        pass
    # Fallback COUNT(*)
    q = f"SELECT COUNT(*) AS c FROM [{table_name}]"
    rs = db.OpenRecordset(q, DB_OPEN_SNAPSHOT)
    try:
        rs.MoveFirst()
    except Exception:
        pass
    rc = int(rs.Fields("c").Value)
    rs.Close()
    return rc

def list_tables_estimates(db, avg_memo_chars=DEFAULT_AVG_MEMO_CHARS, text_fill=DEFAULT_TEXT_FILL):
    results = []
    for tdef in iter_collection(db.TableDefs):
        tname = str(tdef.Name)
        if not is_user_local_table(tdef):
            continue

        notes = []
        try:
            for fld in iter_collection(tdef.Fields):
                ftype = int(fld.Type)
                if ftype == DB_MEMO:
                    notes.append("LongText")
                if ftype in (DB_LONG_BINARY, DB_ATTACHMENT):
                    notes.append("Attachment/OLE")
        except Exception:
            pass

        try:
            rows = fast_row_count(db, tname)
        except Exception as e:
            rows = 0
            notes.append(f"RowCountErr:{type(e).__name__}")

        try:
            per_row = estimated_row_bytes(tdef, avg_memo_chars, text_fill)
            est_bytes = int(((per_row + ROW_OVERHEAD_BYTES) * rows) * INDEX_OVERHEAD_MULT)
        except Exception as e:
            est_bytes = 0
            notes.append(f"SizeEstErr:{type(e).__name__}")

        results.append({
            "TableName": tname,
            "Rows": rows,
            "EstBytes": est_bytes,
            "EstMB": est_bytes / (1024 * 1024),
            "Notes": ";".join(sorted(set(notes))) if notes else ""
        })
    results.sort(key=lambda r: r["EstMB"], reverse=True)
    return results

def save_csv(results, out_path):
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["TableName", "Rows", "EstMB", "EstBytes", "Notes"])
        for r in results:
            w.writerow([r["TableName"], r["Rows"], f"{r['EstMB']:.3f}", r["EstBytes"], r["Notes"]])

def print_table(results, top_n=None):
    rows = results if top_n is None else results[:top_n]
    if not rows:
        print("No local tables found.")
        return
    name_w = max(9, *(len(r["TableName"]) for r in rows))
    print(f"{'TableName'.ljust(name_w)}  {'Rows':>12}  {'EstMB':>10}  {'Notes'}")
    print("-" * (name_w + 2 + 12 + 2 + 10 + 2 + 20))
    for r in rows:
        print(f"{r['TableName'].ljust(name_w)}  {r['Rows']:>12}  {r['EstMB']:>10.3f}  {r['Notes']}")

# ----------------- Main -----------------

def parse_args():
    frag = None
    path = None
    for arg in sys.argv[1:]:
        if arg.startswith("--path="):
            path = arg.split("=", 1)[1].strip('"').strip("'")
        else:
            frag = arg
    return frag, path

def main():
    frag, path = parse_args()
    if frag is None and path is None:
        frag = "IMA_CGME"  # default convenience

    app, db, full_path, base = pick_database(name_fragment=(frag or ""), explicit_path=path)
    print(f"Using Access {getattr(app, 'Version', '?.?')} | DB: {full_path}")

    results = list_tables_estimates(db)
    out_dir = os.path.dirname(full_path)
    out_base = os.path.splitext(base)[0]
    out_csv = os.path.join(out_dir, f"{out_base}_table_sizes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    save_csv(results, out_csv)

    print_table(results)
    print(f"\nSaved: {out_csv}")

if __name__ == "__main__":
    main()

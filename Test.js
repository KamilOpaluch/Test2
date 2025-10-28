# access_table_sizes.py
import os
import re
import sys
import csv
from datetime import datetime

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    raise SystemExit("Requires pywin32. Install with: pip install pywin32")

# ---------- Estimation tunables ----------
ROW_OVERHEAD_BYTES = 24
INDEX_OVERHEAD_MULT = 1.10
DEFAULT_AVG_MEMO_CHARS = 200
DEFAULT_TEXT_FILL = 0.5

# ---------- DAO constants ----------
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

ACC_EXT_RX = re.compile(r'(?:"|\')?([A-Za-z]:[^"\']+\.(?:accdb|mdb))(?:"|\')?', re.IGNORECASE)

# ---------- Helpers to open DB without relying on UI attachment ----------

def make_dao_engine():
    """Create a DAO DBEngine (works even if no Access UI is attached)."""
    # Try modern ACE first, then older DAO
    for progid in ("DAO.DBEngine.120", "DAO.DBEngine.36"):
        try:
            return win32.Dispatch(progid)
        except Exception:
            continue
    raise RuntimeError("Could not create DAO.DBEngine (ACE/DAO not installed?).")

def open_db_direct(full_path):
    """Open an Access DB file via DAO directly."""
    full_path = os.path.abspath(full_path)
    engine = make_dao_engine()
    ws = engine.Workspaces(0)
    db = ws.OpenDatabase(full_path)
    _ = db.TableDefs.Count  # sanity
    return db, full_path, os.path.basename(full_path)

# ---------- Try to attach to running Access instances (optional) ----------

def enumerate_access_instances():
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
        # Access.Application moniker
        if "Access.Application" in name or name.lower().endswith("!{73fddc80-aea9-101a-98a7-00aa00374959}"):
            try:
                obj = rot.GetObject(mk)
                apps.append(win32.Dispatch(obj))
            except Exception:
                pass
    return apps

def try_get_active_db_from_app(app):
    """Return (db, full_path, base_name) or (None, '', '')."""
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

def find_running_access_db_by_fragment(fragment):
    """Attach to a running Access window whose open file matches the fragment."""
    apps = enumerate_access_instances()
    if not apps:
        return None  # no UI instances visible via ROT
    candidates = []
    for app in apps:
        db, full, base = try_get_active_db_from_app(app)
        if db and full:
            candidates.append((db, full, base))
    if not candidates:
        return None
    frag = fragment.lower()
    for db, full, base in candidates:
        if frag in base.lower():
            return db, full, base
    # Fallback: first one
    return candidates[0]

# ---------- Fallback: find Access DB via WMI process command lines ----------

def wmi_find_access_paths():
    """Return list of .accdb/.mdb paths found in msaccess.exe command lines."""
    paths = []
    try:
        wmi = win32.GetObject("winmgmts:")
        procs = wmi.ExecQuery("SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name='MSACCESS.EXE'")
        for p in procs:
            cmd = (p.CommandLine or "")
            # Find any .accdb/.mdb in the command line
            matches = ACC_EXT_RX.findall(cmd)
            for m in matches:
                if os.path.isfile(m):
                    paths.append(os.path.abspath(m))
    except Exception:
        pass
    # De-dup while preserving order
    seen = set()
    uniq = []
    for p in paths:
        if p.lower() not in seen:
            uniq.append(p)
            seen.add(p.lower())
    return uniq

def pick_db_path_by_fragment(paths, fragment):
    if not fragment:
        return paths[0] if paths else None
    frag = fragment.lower()
    for p in paths:
        if frag in os.path.basename(p).lower():
            return p
    return paths[0] if paths else None

# ---------- Table size estimation ----------

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
            return False  # linked table
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
    # Try dbOpenTable for instant count
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

# ---------- Arg parsing & main ----------

def parse_args():
    frag = None
    path = None
    for arg in sys.argv[1:]:
        if arg.startswith("--path="):
            path = arg.split("=", 1)[1].strip('"').strip("'")
        else:
            frag = arg
    if frag is None and path is None:
        frag = "IMA_CGME"  # default convenience
    return frag, path

def main():
    pythoncom.CoInitialize()
    try:
        frag, explicit_path = parse_args()

        # 1) If we got a path, just open it directly via DAO
        if explicit_path:
            db, full, base = open_db_direct(explicit_path)
        else:
            # 2) Try to attach to a running Access instance that has a DB open
            attached = find_running_access_db_by_fragment(frag)
            if attached:
                db, full, base = attached
            else:
                # 3) Try to read the DB path from msaccess.exe command lines
                paths = wmi_find_access_paths()
                target = pick_db_path_by_fragment(paths, frag)
                if not target:
                    raise RuntimeError("Could not find an open Access DB via COM or WMI. "
                                       "Run with --path=\"C:\\full\\path\\YourDb.accdb\" "
                                       "or start Access by double-clicking the .accdb so the path "
                                       "appears in the process command line.")
                db, full, base = open_db_direct(target)

        print(f"DB: {full}")
        results = list_tables_estimates(db)
        out_dir = os.path.dirname(full)
        out_base = os.path.splitext(base)[0]
        out_csv = os.path.join(out_dir, f"{out_base}_table_sizes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        save_csv(results, out_csv)
        print_table(results)
        print(f"\nSaved: {out_csv}")

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

if __name__ == "__main__":
    main()
        

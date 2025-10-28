# access_table_sizes.py
import os
import csv
import sys
from datetime import datetime

try:
    import win32com.client as win32
except ImportError:
    raise SystemExit("pywin32 is required. Install with: pip install pywin32")

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

def attach_access():
    try:
        app = win32.GetActiveObject("Access.Application")
        return app
    except Exception as e:
        raise RuntimeError("No running Microsoft Access instance found. Open your .accdb and try again.") from e

def iter_collection(coll):
    cnt = int(coll.Count)
    # Try 0-based
    try:
        for i in range(cnt):
            yield coll.Item(i)
        return
    except Exception:
        pass
    # Fallback 1-based
    for i in range(1, cnt + 1):
        yield coll.Item(i)

def pick_open_database(app, name_fragment=""):
    # Get the active file path
    try:
        full_path = str(app.CurrentProject.FullName)
    except Exception as e:
        raise RuntimeError("Attached Access instance has no database open.") from e

    base = os.path.basename(full_path)
    if not base:
        raise RuntimeError("Could not resolve active database path from Access.")

    if name_fragment and name_fragment.lower() not in base.lower():
        print(f"[warn] Active Access DB is '{base}', not matching fragment '{name_fragment}'. Proceeding with the active DB.")

    # 1) Try CurrentDb()
    db = None
    try:
        db = app.CurrentDb()
        # sanity access
        _ = db.TableDefs.Count
        return db, full_path, base
    except Exception:
        db = None

    # 2) Fallback: open via the same Access instance's DAO DBEngine
    try:
        ws0 = app.DBEngine.Workspaces(0)
        db = ws0.OpenDatabase(full_path)
        # sanity access
        _ = db.TableDefs.Count
        return db, full_path, base
    except Exception as e:
        raise RuntimeError("Failed to obtain a DAO.Database handle for the active file.") from e

def is_user_local_table(tdef):
    name = str(tdef.Name)
    if name.startswith("MSys"):
        return False
    try:
        if len(str(tdef.Connect)) > 0:
            return False  # linked table: lives elsewhere
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
        return 0  # unknown, flagged elsewhere
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
    # Fallback: COUNT(*)
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

def main(name_fragment="IMA_CGME", avg_memo_chars=DEFAULT_AVG_MEMO_CHARS, text_fill=DEFAULT_TEXT_FILL):
    app = attach_access()
    db, full_path, base = pick_open_database(app, name_fragment)

    print(f"Attached to Access {getattr(app, 'Version', '?.?')} | Active DB: {full_path}")

    results = list_tables_estimates(db, avg_memo_chars=avg_memo_chars, text_fill=text_fill)

    out_dir = os.path.dirname(full_path)
    out_base = os.path.splitext(base)[0]
    out_csv = os.path.join(out_dir, f"{out_base}_table_sizes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    save_csv(results, out_csv)

    print_table(results)
    print(f"\nSaved: {out_csv}")

if __name__ == "__main__":
    fragment = sys.argv[1] if len(sys.argv) > 1 else "IMA_CGME"
    main(fragment)

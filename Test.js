import os
import csv
import math
import sys
from datetime import datetime

try:
    import win32com.client as win32
except ImportError:
    raise SystemExit("pywin32 is required. Install with: pip install pywin32")

# ---- Tunables for the estimator ----
ROW_OVERHEAD_BYTES = 24              # rough per-row overhead
INDEX_OVERHEAD_MULT = 1.10           # add ~10% for indexes
DEFAULT_AVG_MEMO_CHARS = 200         # assumed avg length for Long Text
DEFAULT_TEXT_FILL = 0.5              # fraction of Short Text max length used

# DAO field type constants (subset)
DB_BOOLEAN = 1
DB_BYTE = 2
DB_INTEGER = 3
DB_LONG = 4
DB_CURRENCY = 5
DB_SINGLE = 6
DB_DOUBLE = 7
DB_DATE = 8
DB_BINARY = 9
DB_TEXT = 10          # Short Text
DB_LONG_BINARY = 11   # OLE Object
DB_MEMO = 12          # Long Text
DB_GUID = 15
DB_DECIMAL = 20
DB_ATTACHMENT = 101   # Attachment (ACE)

# DAO OpenRecordset types
DB_OPEN_TABLE = 1
DB_OPEN_SNAPSHOT = 4

def attach_access():
    """Attach to the already-running Access.Application."""
    try:
        return win32.GetActiveObject("Access.Application")
    except Exception as e:
        raise RuntimeError("No running Microsoft Access instance found. Open your database first.") from e

def iter_collection(coll):
    """Robustly iterate a COM collection that may be 0- or 1-based."""
    count = int(coll.Count)
    # Try 0-based first
    try:
        for i in range(count):
            yield coll.Item(i)
        return
    except Exception:
        pass
    # Fallback to 1-based
    for i in range(1, count + 1):
        yield coll.Item(i)

def pick_open_database(app, name_fragment):
    """Find an open DAO.Database whose filename contains name_fragment."""
    dbs = app.DBEngine.Workspaces(0).Databases
    for db in iter_collection(dbs):
        base = os.path.basename(str(db.Name))
        if name_fragment.lower() in base.lower():
            return db
    raise RuntimeError(f"No open database whose filename contains '{name_fragment}' was found.")

def is_user_local_table(tdef):
    """Local, non-system, non-linked table?"""
    name = str(tdef.Name)
    if name.startswith("MSys"):
        return False
    # Linked if Connect string present
    try:
        if len(str(tdef.Connect)) > 0:
            return False
    except Exception:
        pass
    return True

def field_size_bytes(fld, avg_memo_chars=DEFAULT_AVG_MEMO_CHARS, text_fill=DEFAULT_TEXT_FILL):
    """Estimated per-row storage for a single field."""
    ftype = int(fld.Type)
    # Some types expose fld.Size (e.g., Short Text max length)
    fsize = 0
    try:
        fsize = int(fld.Size)
    except Exception:
        pass

    if ftype == DB_BOOLEAN:        return 1
    if ftype == DB_BYTE:           return 1
    if ftype == DB_INTEGER:        return 2
    if ftype == DB_LONG:           return 4
    if ftype == DB_SINGLE:         return 4
    if ftype == DB_DOUBLE:         return 8
    if ftype == DB_CURRENCY:       return 8
    if ftype == DB_DECIMAL:        return 12
    if ftype == DB_DATE:           return 8
    if ftype == DB_GUID:           return 16
    if ftype == DB_BINARY:         return max(fsize, 0)
    if ftype == DB_TEXT:           return int(round((fsize * 2) * text_fill))  # Unicode chars
    if ftype == DB_MEMO:           return int(avg_memo_chars * 2)              # Long Text: avg guess
    if ftype in (DB_LONG_BINARY, DB_ATTACHMENT):
        return 0  # unknown (can be huge); we’ll flag in Notes
    # Unknown/Calculated/etc.
    return 0

def estimated_row_bytes(tdef, avg_memo_chars, text_fill):
    total = 0
    for fld in iter_collection(tdef.Fields):
        total += field_size_bytes(fld, avg_memo_chars, text_fill)
    return total

def fast_row_count(db, table_name):
    """Try table-type recordset for instant count; fallback to COUNT(*)."""
    # First try dbOpenTable (instant for native Access tables)
    try:
        rs = db.OpenRecordset(table_name, DB_OPEN_TABLE)
        rc = int(rs.RecordCount)
        rs.Close()
        return rc
    except Exception:
        pass
    # Fallback to COUNT(*)
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
        # pre-flag if any Attachment/OLE/LongText present
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
    # Sort by estimated MB desc
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
    # column widths
    name_w = max(9, *(len(r["TableName"]) for r in rows)) if rows else 9
    print(f"{'TableName'.ljust(name_w)}  {'Rows':>12}  {'EstMB':>10}  {'Notes'}")
    print("-" * (name_w + 2 + 12 + 2 + 10 + 2 + 20))
    for r in rows:
        print(f"{r['TableName'].ljust(name_w)}  {r['Rows']:>12}  {r['EstMB']:>10.3f}  {r['Notes']}")

def main(name_fragment="IMA_CGME", avg_memo_chars=DEFAULT_AVG_MEMO_CHARS, text_fill=DEFAULT_TEXT_FILL):
    app = attach_access()
    db = pick_open_database(app, name_fragment)

    results = list_tables_estimates(db, avg_memo_chars=avg_memo_chars, text_fill=text_fill)

    base = os.path.splitext(os.path.basename(str(db.Name)))[0]
    out_csv = os.path.join(os.path.dirname(str(db.Name)),
                           f"{base}_table_sizes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    save_csv(results, out_csv)

    print(f"Database: {db.Name}")
    print_table(results)
    print(f"\nSaved: {out_csv}")

if __name__ == "__main__":
    # Allow optional name fragment via CLI: python access_table_sizes.py IMA_CGME
    fragment = sys.argv[1] if len(sys.argv) > 1 else "IMA_CGME"
    main(fragment)

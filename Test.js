import os, sys, csv, re
from datetime import datetime

# Optional GUI picker if no path is provided
def pick_file_dialog():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        path = filedialog.askopenfilename(
            title="Select Access database",
            filetypes=[("Access DB", "*.accdb *.mdb"), ("All files","*.*")]
        )
        return path or None
    except Exception:
        return None

# --- Estimation knobs ---
ROW_OVERHEAD_BYTES = 24
INDEX_OVERHEAD_MULT = 1.10
DEFAULT_AVG_MEMO_CHARS = 200
DEFAULT_TEXT_FILL = 0.5

def parse_args():
    frag = None
    path = None
    for arg in sys.argv[1:]:
        if arg.startswith("--path="):
            path = arg.split("=",1)[1].strip("'").strip('"')
        else:
            frag = arg
    return frag, path

# ===== DAO path (best: can detect linked tables) =====
def try_open_dao(full_path):
    import win32com.client as win32
    for progid in ("DAO.DBEngine.120", "DAO.DBEngine.36"):
        try:
            engine = win32.Dispatch(progid)
            ws = engine.Workspaces(0)
            db = ws.OpenDatabase(full_path)
            _ = db.TableDefs.Count
            return db
        except Exception:
            continue
    return None

def dao_iter(coll):
    cnt = int(coll.Count)
    try:
        for i in range(cnt):
            yield coll.Item(i); return
    except Exception:
        pass
    for i in range(1, cnt+1):
        yield coll.Item(i)

def dao_is_local(tdef):
    nm = str(tdef.Name)
    if nm.startswith("MSys"): return False
    try:
        if len(str(tdef.Connect)) > 0:  # linked
            return False
    except Exception:
        pass
    return True

# DAO field sizes
DB_BOOLEAN,DB_BYTE,DB_INTEGER,DB_LONG,DB_CURRENCY,DB_SINGLE,DB_DOUBLE,DB_DATE,DB_BINARY,DB_TEXT,DB_LONG_BINARY,DB_MEMO,DB_GUID,DB_DECIMAL,DB_ATTACHMENT = \
    1,2,3,4,5,6,7,8,9,10,11,12,15,20,101

def dao_field_bytes(fld, avg_memo=DEFAULT_AVG_MEMO_CHARS, fill=DEFAULT_TEXT_FILL):
    t = int(fld.Type)
    try: sz = int(fld.Size)
    except: sz = 0
    if t==DB_BOOLEAN or t==DB_BYTE: return 1
    if t==DB_INTEGER: return 2
    if t==DB_LONG or t==DB_SINGLE: return 4
    if t==DB_DOUBLE or t==DB_CURRENCY: return 8
    if t==DB_DECIMAL: return 12
    if t==DB_DATE: return 8
    if t==DB_GUID: return 16
    if t==DB_BINARY: return max(sz,0)
    if t==DB_TEXT: return int(round((sz*2)*fill))
    if t==DB_MEMO: return int(avg_memo*2)
    if t in (DB_LONG_BINARY, DB_ATTACHMENT): return 0
    return 0

def dao_row_bytes(tdef):
    total = 0
    for fld in dao_iter(tdef.Fields):
        total += dao_field_bytes(fld)
    return total

def dao_rowcount(db, tname):
    DB_OPEN_TABLE, DB_OPEN_SNAPSHOT = 1, 4
    try:
        rs = db.OpenRecordset(tname, DB_OPEN_TABLE)
        rc = int(rs.RecordCount); rs.Close(); return rc
    except Exception:
        rs = db.OpenRecordset(f"SELECT COUNT(*) AS c FROM [{tname}]", DB_OPEN_SNAPSHOT)
        try: rs.MoveFirst()
        except: pass
        rc = int(rs.Fields("c").Value); rs.Close(); return rc

def analyze_with_dao(full_path):
    db = try_open_dao(full_path)
    if not db: return None  # fall back to ODBC
    results = []
    for tdef in dao_iter(db.TableDefs):
        name = str(tdef.Name)
        if not dao_is_local(tdef): continue
        notes = []
        # flag heavy types
        try:
            for fld in dao_iter(tdef.Fields):
                if int(fld.Type)==DB_MEMO: notes.append("LongText")
                if int(fld.Type) in (DB_LONG_BINARY, DB_ATTACHMENT): notes.append("Attachment/OLE")
        except Exception:
            pass
        try:
            rows = dao_rowcount(db, name)
            per_row = dao_row_bytes(tdef)
            est = int(((per_row + ROW_OVERHEAD_BYTES)*rows) * INDEX_OVERHEAD_MULT)
        except Exception as e:
            rows, est = 0, 0
            notes.append(f"Err:{type(e).__name__}")
        results.append({
            "TableName": name,
            "Rows": rows,
            "EstBytes": est,
            "EstMB": est/(1024*1024),
            "Notes": ";".join(sorted(set(notes))) if notes else ""
        })
    results.sort(key=lambda r: r["EstMB"], reverse=True)
    return results

# ===== ODBC path (fallback; may include linked tables) =====
def try_open_odbc(full_path):
    import pyodbc
    conn = pyodbc.connect(
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="+full_path+";",
        autocommit=True
    )
    return conn

def odbc_type_bytes(type_name, column_size, avg_memo=DEFAULT_AVG_MEMO_CHARS, fill=DEFAULT_TEXT_FILL):
    t = (type_name or "").upper()
    cs = int(column_size) if column_size else 0
    # Map common Access ODBC names
    if t in ("BIT","YESNO"): return 1
    if t in ("BYTE"): return 1
    if t in ("SHORT","SMALLINT"): return 2
    if t in ("LONG","INTEGER","INT"): return 4
    if t in ("SINGLE","REAL"): return 4
    if t in ("DOUBLE","FLOAT","NUMBER"): return 8
    if t in ("CURRENCY","MONEY","DECIMAL","NUMERIC"): return 8
    if t in ("DATETIME","DATE","TIME","TIMESTAMP"): return 8
    if t in ("GUID","UNIQUEIDENTIFIER"): return 16
    if t in ("CHAR","NCHAR","VARCHAR","NVARCHAR","TEXT"):
        return int(round((max(cs,1) * 2) * fill))
    if t in ("LONGCHAR","MEMO","NTEXT"):
        return int(avg_memo * 2)
    if t in ("BINARY","VARBINARY","LONGBINARY","IMAGE","OLEOBJECT"):
        return 0
    return 0

def odbc_list_user_tables(conn):
    # Best-effort: filter out system tables by name; we can't always know links via ODBC
    cur = conn.cursor()
    names = []
    for row in cur.tables():
        if row.table_type and row.table_type.upper() not in ("TABLE","VIEW"):
            continue
        name = row.table_name
        if not name or name.startswith("MSys") or name.startswith("~TMP"):  # skip system/tmp
            continue
        names.append(name)
    # Try MSysObjects to spot links (may fail if perms don’t allow)
    linked = set()
    try:
        cur.execute("""SELECT Name, Type FROM MSysObjects
                       WHERE Type IN (1,4,6) AND Left(Name,1)<>'~' AND Left(Name,4)<>'MSys'""")
        for n,t in cur.fetchall():
            # 1=local, 4/6=linked (varies by kind)
            if t in (4,6): linked.add(n)
    except Exception:
        pass
    return names, linked

def analyze_with_odbc(full_path):
    conn = try_open_odbc(full_path)
    cur = conn.cursor()
    names, linked = odbc_list_user_tables(conn)
    results = []
    for name in names:
        notes = []
        if name in linked:
            notes.append("Linked?")  # best guess
        # columns metadata for per-row estimate
        per_row = 0
        try:
            for col in cur.columns(table=name):
                per_row += odbc_type_bytes(col.type_name, col.column_size)
        except Exception:
            pass
        # row count
        rows = 0
        try:
            cur.execute(f"SELECT COUNT(*) FROM [{name}]")
            rows = int(cur.fetchone()[0])
        except Exception as e:
            notes.append(f"RowCountErr:{type(e).__name__}")
        est = int(((per_row + ROW_OVERHEAD_BYTES)*rows) * INDEX_OVERHEAD_MULT)
        results.append({
            "TableName": name,
            "Rows": rows,
            "EstBytes": est,
            "EstMB": est/(1024*1024),
            "Notes": ";".join(sorted(set(notes))) if notes else ""
        })
    results.sort(key=lambda r: r["EstMB"], reverse=True)
    return results

def save_csv(results, out_path):
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["TableName","Rows","EstMB","EstBytes","Notes"])
        for r in results:
            w.writerow([r["TableName"], r["Rows"], f"{r['EstMB']:.3f}", r["EstBytes"], r["Notes"]])

def print_table(results):
    if not results:
        print("No tables found."); return
    name_w = max(9, *(len(r["TableName"]) for r in results))
    print(f"{'TableName'.ljust(name_w)}  {'Rows':>12}  {'EstMB':>10}  Notes")
    print("-"*(name_w+2+12+2+10+2+20))
    for r in results:
        print(f"{r['TableName'].ljust(name_w)}  {r['Rows']:>12}  {r['EstMB']:>10.3f}  {r['Notes']}")

def main():
    frag, path = parse_args()
    if not path:
        # if you passed a fragment, we still need a file; show picker
        path = pick_file_dialog()
    if not path:
        print("Please rerun with --path=\"C:\\path\\to\\YourDB.accdb\" (or use the picker).")
        return
    path = os.path.abspath(path)
    base = os.path.basename(path)
    out_csv = os.path.join(os.path.dirname(path),
                           f"{os.path.splitext(base)[0]}_table_sizes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

    # Try DAO first (best for excluding linked)
    results = analyze_with_dao(path)
    if results is None:
        print("[info] DAO not available or mismatched bitness; using ODBC fallback.")
        results = analyze_with_odbc(path)

    print(f"DB: {path}")
    print_table(results)
    save_csv(results, out_csv)
    print(f"\nSaved: {out_csv}")

if __name__ == "__main__":
    main()

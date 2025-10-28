import os, sys, csv, time, platform

# ===== Pretty logging =====
def log(msg): print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)

# ===== Estimation knobs =====
ROW_OVERHEAD_BYTES = 24
INDEX_OVERHEAD_MULT = 1.10
DEFAULT_AVG_MEMO_CHARS = 200
DEFAULT_TEXT_FILL = 0.5

# ===== DAO constants =====
DB_BOOLEAN,DB_BYTE,DB_INTEGER,DB_LONG,DB_CURRENCY,DB_SINGLE,DB_DOUBLE,DB_DATE,DB_BINARY,DB_TEXT,DB_LONG_BINARY,DB_MEMO,DB_GUID,DB_DECIMAL,DB_ATTACHMENT = \
    1,2,3,4,5,6,7,8,9,10,11,12,15,20,101
DB_OPEN_TABLE, DB_OPEN_SNAPSHOT = 1, 4

def iter_coll(coll):
    cnt = int(coll.Count)
    try:
        for i in range(cnt):
            yield coll.Item(i)
        return
    except Exception:
        pass
    for i in range(1, cnt+1):
        yield coll.Item(i)

def is_local_table(tdef):
    nm = str(tdef.Name)
    if nm.startswith("MSys"): return False
    try:
        if len(str(tdef.Connect)) > 0: return False
    except Exception:
        pass
    return True

def field_bytes(fld, avg_memo=DEFAULT_AVG_MEMO_CHARS, fill=DEFAULT_TEXT_FILL):
    t = int(fld.Type)
    try: sz = int(fld.Size)
    except: sz = 0
    if t in (DB_BOOLEAN, DB_BYTE): return 1
    if t == DB_INTEGER: return 2
    if t in (DB_LONG, DB_SINGLE): return 4
    if t in (DB_DOUBLE, DB_CURRENCY): return 8
    if t == DB_DECIMAL: return 12
    if t == DB_DATE: return 8
    if t == DB_GUID: return 16
    if t == DB_BINARY: return max(sz,0)
    if t == DB_TEXT: return int(round((sz*2)*fill))
    if t == DB_MEMO: return int(avg_memo*2)
    if t in (DB_LONG_BINARY, DB_ATTACHMENT): return 0
    return 0

def row_bytes(tdef):
    return sum(field_bytes(f) for f in iter_coll(tdef.Fields))

def fast_rowcount(db, tname):
    try:
        rs = db.OpenRecordset(tname, DB_OPEN_TABLE)
        rc = int(rs.RecordCount); rs.Close(); return rc
    except Exception:
        rs = db.OpenRecordset(f"SELECT COUNT(*) AS c FROM [{tname}]", DB_OPEN_SNAPSHOT)
        try: rs.MoveFirst()
        except: pass
        rc = int(rs.Fields("c").Value); rs.Close(); return rc

def analyze_currentdb(app):
    db = app.CurrentDb()  # live handle to the open session's DB
    full = str(app.CurrentProject.FullName)
    base = os.path.basename(full) if full else "(unknown)"

    log(f"Active window DB path reported by Access: {full or '(empty)'}")
    # Even if FullName is empty (rare), CurrentDb still works for TableDefs/queries.

    tables = [t for t in iter_coll(db.TableDefs) if is_local_table(t)]
    log(f"Local tables found: {len(tables)}")

    results = []
    for i, tdef in enumerate(tables, 1):
        tname = str(tdef.Name)
        log(f"[{i}/{len(tables)}] {tname}: estimating row width…")
        per = row_bytes(tdef)
        log(f" -> RowWidthBytes ≈ {per}")
        log(" -> Counting rows (fast)…")
        rows = fast_rowcount(db, tname)
        est = int(((per + ROW_OVERHEAD_BYTES)*rows) * INDEX_OVERHEAD_MULT)
        results.append({
            "TableName": tname,
            "Rows": rows,
            "EstBytes": est,
            "EstMB": est/(1024*1024),
            "Notes": ""
        })
    results.sort(key=lambda r: r["EstMB"], reverse=True)
    return results, full, base

def save_csv(results, suggested_dir, base_hint):
    out_dir = suggested_dir if suggested_dir and os.path.isdir(suggested_dir) else os.getcwd()
    stem = os.path.splitext(base_hint or 'AccessDb')[0]
    out_csv = os.path.join(out_dir, f"{stem}_table_sizes_{time.strftime('%Y%m%d_%H%M%S')}.csv")
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["TableName","Rows","EstMB","EstBytes","Notes"])
        for r in results:
            w.writerow([r["TableName"], r["Rows"], f"{r['EstMB']:.3f}", r["EstBytes"], r["Notes"]])
    return out_csv

def main():
    log(f"Python {platform.python_version()} | arch={platform.architecture()[0]} | user={os.getlogin()}")
    try:
        import pythoncom, win32com.client as win32
    except Exception as e:
        log("ERROR: pywin32 is required. Install with: pip install pywin32")
        sys.exit(1)

    pythoncom.CoInitialize()
    try:
        app = None
        # Try GetObject(None, 'Access.Application')
        try:
            log("Attach attempt #1: GetObject(None, 'Access.Application') …")
            app = win32.GetObject(None, "Access.Application")
            log(" -> Success.")
        except Exception as e1:
            log(f" -> Failed: {type(e1).__name__}: {e1}")
            # Try GetActiveObject
            try:
                log("Attach attempt #2: GetActiveObject('Access.Application') …")
                app = win32.GetActiveObject("Access.Application")
                log(" -> Success.")
            except Exception as e2:
                log(f" -> Failed: {type(e2).__name__}: {e2}")
                # Enumerate ROT manually
                try:
                    log("Attach attempt #3: Enumerating ROT for Access.Application …")
                    ctx = pythoncom.CreateBindCtx(0)
                    rot = pythoncom.GetRunningObjectTable()
                    enum = rot.EnumRunning()
                    found = None
                    while True:
                        fetched = enum.Next(1)
                        if not fetched: break
                        mk = fetched[0]
                        try:
                            name = mk.GetDisplayName(ctx, None)
                        except pythoncom.com_error:
                            continue
                        if ("Access.Application" in name) or name.lower().endswith("!{73fddc80-aea9-101a-98a7-00aa00374959}"):
                            try:
                                obj = rot.GetObject(mk)
                                candidate = win32.Dispatch(obj)
                                # Sanity: only accept if it has a CurrentDb
                                _ = candidate.CurrentDb()
                                found = candidate
                                break
                            except Exception:
                                pass
                    if found is None:
                        log(" -> ROT scan found no attachable Access instance.")
                    else:
                        app = found
                        log(" -> Success (ROT).")
                except Exception as e3:
                    log(f" -> ROT enumeration failed: {type(e3).__name__}: {e3}")

        if app is None:
            log("")
            log("❌ Could not attach to the already-open Access session.")
            log("Most common reasons (and fixes):")
            log(" - Bitness mismatch (Access 32-bit vs Python 64-bit or vice-versa) → run Python matching Access bitness.")
            log(" - UAC elevation mismatch (one is Admin, the other isn’t) → run both at the same elevation.")
            log(" - Different user desktop/session (RDP/Service) → run under the same interactive user session.")
            log(" - Click-to-Run isolation blocking ROT → try starting Python from the same user context (non-admin).")
            sys.exit(2)

        # Good: attached
        try:
            ver = getattr(app, "Version", "?.?")
        except Exception:
            ver = "?.?"
        log(f"✅ Attached to Access {ver}. Using the LIVE CurrentDb() from that window.")
        results, full, base = analyze_currentdb(app)

        out_csv = save_csv(results, os.path.dirname(full) if full else None, base or "AccessDb")
        # Print summary
        if results:
            width = max(9, *(len(r["TableName"]) for r in results[:20]))
            print(f"\n{'TableName'.ljust(width)}  {'Rows':>12}  {'EstMB':>10}")
            print("-"*(width+2+12+2+10))
            for r in results[:20]:
                print(f"{r['TableName'].ljust(width)}  {r['Rows']:>12}  {r['EstMB']:>10.3f}")
            if len(results) > 20:
                print(f"... and {len(results)-20} more")
        log(f"CSV saved to: {out_csv}")

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

if __name__ == "__main__":
    main()

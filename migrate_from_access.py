"""
Migrate the Product Cost Management Access database into a local SQLite file.
Re-runnable: drops and recreates the data tables on every run.
Includes a highly stable live terminal progress bar (tqdm).
"""
from __future__ import annotations

import os
import re
import sqlite3
import sys
import unicodedata

import win32com.client  # type: ignore  (from pywin32)
from tqdm import tqdm   # Fortschrittsbalken-Bibliothek

# --- configuration ----------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ACCDB = os.path.join(os.path.dirname(BASE_DIR), "Product_Cost Management_Project_DB.accdb")
SQLITE_PATH = os.path.join(BASE_DIR, "data.sqlite")

LOCAL_TABLES = [
    "tbl_alle_anzeigen",
    "tbl_alle_anzeigen_einf_übersicht",
    "tbl_alle_anzeigen_vm",
    "tbl_summe_demand",
]

LINKED_TABLES = [
    "tbl_CS15M_BOM_analysis",
]

ROWID_TABLE = "tbl_alle_anzeigen"

DAO_TYPE_TO_SQLITE = {
    1: "INTEGER", 2: "INTEGER", 3: "INTEGER", 4: "INTEGER",
    5: "REAL", 6: "REAL", 7: "REAL", 8: "TEXT",
    10: "TEXT", 12: "TEXT", 15: "TEXT", 16: "INTEGER",
    18: "TEXT", 19: "REAL", 20: "REAL",
}

_UMLAUT = {"ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss", "Ä": "Ae", "Ö": "Oe", "Ü": "Ue"}

def safe_name(name: str) -> str:
    for src, dst in _UMLAUT.items():
        name = name.replace(src, dst)
    name = "".join(c for c in unicodedata.normalize("NFKD", name) if not unicodedata.combining(c))
    name = re.sub(r"[^0-9A-Za-z]+", "_", name).strip("_")
    if not name: name = "col"
    if name[0].isdigit(): name = "c_" + name
    return name

def unique_safe_names(originals: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    seen: set[str] = set()
    for orig in originals:
        base = safe_name(orig)
        cand = base
        i = 1
        while cand.lower() in seen:
            i += 1
            cand = f"{base}_{i}"
        seen.add(cand.lower())
        out[orig] = cand
    return out

def to_py(value):
    if value is None: return None
    try:
        import datetime
        if isinstance(value, datetime.datetime): return value.isoformat(sep=" ")
        if isinstance(value, datetime.date): return value.isoformat()
    except Exception: pass
    if isinstance(value, (int, float, str, bytes)): return value
    return str(value)

def read_table(db, table_name):
    """Liest die Access-Tabelle in extrem schnellen 5.000er Blöcken aus."""
    # Zeilenanzahl holen
    rs_count = db.OpenRecordset(f"SELECT COUNT(*) FROM [{table_name}]")
    total_rows = rs_count.Fields.Item(0).Value
    rs_count.Close()

    rs = db.OpenRecordset(f"SELECT * FROM [{table_name}]")
    fields_obj = rs.Fields
    fields_count = fields_obj.Count
    
    originals = [fields_obj.Item(i).Name for i in range(fields_count)]
    types = [fields_obj.Item(i).Type for i in range(fields_count)]
    
    rows = []
    if not rs.EOF: 
        rs.MoveFirst()
    
    # In 5.000er Blöcken statt einzeln lesen (GetRows)
    CHUNK_SIZE = 5000
    
    with tqdm(total=total_rows, desc=f"  Lese {table_name[:25]:25s}", unit="zeilen") as pbar:
        while not rs.EOF:
            # Holt bis zu 5000 Zeilen auf einmal in den Arbeitsspeicher
            data_chunk = rs.GetRows(CHUNK_SIZE)
            if not data_chunk:
                break
                
            # Transponieren und Werte konvertieren
            num_cols = len(data_chunk)
            num_rows = len(data_chunk[0])
            
            for r_idx in range(num_rows):
                row_data = [to_py(data_chunk[c_idx][r_idx]) for c_idx in range(num_cols)]
                rows.append(row_data)
            
            pbar.update(num_rows)
            
    rs.Close()
    return originals, types, rows

def create_and_fill(con, table_name, originals, types, rows, add_rowid):
    cur = con.cursor()
    safe_tbl = safe_name(table_name)
    safe_map = unique_safe_names(originals)
    col_defs = []
    if add_rowid:
        col_defs.append('"_rowid" INTEGER PRIMARY KEY AUTOINCREMENT')
    for orig, dao_t in zip(originals, types):
        aff = DAO_TYPE_TO_SQLITE.get(dao_t, "TEXT")
        col_defs.append(f'"{safe_map[orig]}" {aff}')

    cur.execute(f'DROP TABLE IF EXISTS "{safe_tbl}"')
    cur.execute(f'CREATE TABLE "{safe_tbl}" ({", ".join(col_defs)})')

    placeholders = ", ".join("?" for _ in originals)
    quoted_cols = ", ".join(f'"{safe_map[o]}"' for o in originals)
    
    cur.executemany(f'INSERT INTO "{safe_tbl}" ({quoted_cols}) VALUES ({placeholders})', rows)

    cur.executemany(
        'INSERT INTO _column_map (tbl, ordinal, safe, original) VALUES (?, ?, ?, ?)',
        [(safe_tbl, i, safe_map[o], o) for i, o in enumerate(originals)],
    )
    con.commit()
    return len(rows)

def main():
    accdb = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_ACCDB
    if not os.path.exists(accdb):
        accdb = r"E:\projects\Cost_Management_Project\Product_Cost Management_Project_DB.accdb"
        if not os.path.exists(accdb):
            accdb = r"c:\temp\Product_Cost Management_Project_DB.accdb"
            if not os.path.exists(accdb):
                sys.exit(f"Access file not found: {accdb}")

    print(f"Source : {accdb}")
    print(f"Target : {SQLITE_PATH}\n")

    con = sqlite3.connect(SQLITE_PATH)
    con.execute("DROP TABLE IF EXISTS _column_map")
    con.execute("CREATE TABLE _column_map (tbl TEXT, ordinal INTEGER, safe TEXT, original TEXT)")
    con.commit()

    app = win32com.client.Dispatch("Access.Application")
    try:
        app.OpenCurrentDatabase(accdb)
        db = app.CurrentDb()

        for t in LOCAL_TABLES:
            originals, types, rows = read_table(db, t)
            n = create_and_fill(con, t, originals, types, rows, add_rowid=(t == ROWID_TABLE))
            print(f"  -> {t} erfolgreich in SQLite gesichert ({n} Zeilen).\n")

        for t in LINKED_TABLES:
            try:
                originals, types, rows = read_table(db, t)
                n = create_and_fill(con, t, originals, types, rows, add_rowid=False)
                print(f"  -> {t} erfolgreich in SQLite gesichert ({n} Zeilen).\n")
            except Exception as exc:
                safe_t = safe_name(t)
                con.execute(f'DROP TABLE IF EXISTS "{safe_t}"')
                con.execute(f'CREATE TABLE "{safe_t}" (Material TEXT)')
                con.commit()
                print(f"  ⚠️ Warnung: Netzwerk-Tabelle {t} nicht erreichbar ({exc})\n")

        con.execute('CREATE INDEX IF NOT EXISTS ix_aa_status ON tbl_alle_anzeigen ("Status")')
        con.commit()
    finally:
        try: app.CloseCurrentDatabase()
        except Exception: pass
        app.Quit()

    con.close()
    print("Done.")

if __name__ == "__main__":
    main()

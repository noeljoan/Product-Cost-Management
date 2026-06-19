"""SQLite access layer for the cost-management web app."""
from __future__ import annotations

import os
import sqlite3
from functools import lru_cache

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data.sqlite")

MAIN_TABLE = "tbl_alle_anzeigen"


def get_conn() -> sqlite3.Connection:
    if not os.path.exists(DB_PATH):
        raise RuntimeError(
            f"{DB_PATH} not found. Run 'py migrate_from_access.py' first."
        )
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


@lru_cache(maxsize=None)
def column_labels(tbl: str = MAIN_TABLE) -> dict[str, str]:
    """safe_name -> original Access label, for nice headers."""
    con = get_conn()
    try:
        rows = con.execute(
            "SELECT safe, original FROM _column_map WHERE tbl = ? ORDER BY ordinal",
            (tbl,),
        ).fetchall()
    finally:
        con.close()
    return {r["safe"]: r["original"] for r in rows}


def table_columns(con: sqlite3.Connection, tbl: str) -> list[str]:
    return [r["name"] for r in con.execute(f'PRAGMA table_info("{tbl}")')]


# --- main grid --------------------------------------------------------------

# Columns shown in the grid (subset of the 35); the detail panel shows everything.
GRID_COLUMNS = [
    "Material", "Materialkurztext", "Werk", "Dispositionselement",
    "Lieferant", "SummevonZugang_Bedarf", "VERPR", "VERPR_vm", "Diff",
    "Datentyp", "Status",
]

# Columns allowed as text filters (LIKE) from the UI.
FILTER_TEXT = {"Material", "Werk", "Lieferant", "Datentyp", "Dispositionselement"}

SORTABLE = set(GRID_COLUMNS)


def fetch_rows(
    *,
    filters: dict[str, str] | None = None,
    status: str | None = None,
    sort: str | None = None,
    direction: str = "asc",
    limit: int = 100,
    offset: int = 0,
) -> tuple[list[sqlite3.Row], int]:
    """Return (rows, total_count) for the main grid with filters/sort/paging."""
    filters = filters or {}
    where: list[str] = []
    params: list = []

    for col, val in filters.items():
        if col in FILTER_TEXT and val:
            where.append(f'"{col}" LIKE ?')
            params.append(f"%{val}%")

    if status in ("0", "1"):
        where.append('CAST("Status" AS TEXT) = ?')
        params.append(status)

    where_sql = (" WHERE " + " AND ".join(where)) if where else ""

    order_sql = ""
    if sort in SORTABLE:
        dir_sql = "DESC" if direction.lower() == "desc" else "ASC"
        order_sql = f' ORDER BY "{sort}" {dir_sql}'
    else:
        order_sql = ' ORDER BY "_rowid" ASC'

    con = get_conn()
    try:
        total = con.execute(
            f'SELECT COUNT(*) AS n FROM "{MAIN_TABLE}"{where_sql}', params
        ).fetchone()["n"]
        rows = con.execute(
            f'SELECT * FROM "{MAIN_TABLE}"{where_sql}{order_sql} LIMIT ? OFFSET ?',
            [*params, limit, offset],
        ).fetchall()
    finally:
        con.close()
    return rows, total


def fetch_one(rowid: int) -> sqlite3.Row | None:
    con = get_conn()
    try:
        return con.execute(
            f'SELECT * FROM "{MAIN_TABLE}" WHERE "_rowid" = ?', (rowid,)
        ).fetchone()
    finally:
        con.close()


# --- status editing ---------------------------------------------------------

def set_status(rowid: int, value: int) -> None:
    con = get_conn()
    try:
        con.execute(
            f'UPDATE "{MAIN_TABLE}" SET "Status" = ? WHERE "_rowid" = ?',
            (str(value), rowid),
        )
        con.commit()
    finally:
        con.close()


def reset_all_status() -> int:
    """Port of the Access 'Status_auf_0_setzen' action query."""
    con = get_conn()
    try:
        cur = con.execute(f'UPDATE "{MAIN_TABLE}" SET "Status" = ?', ("0",))
        con.commit()
        return cur.rowcount
    finally:
        con.close()


def marked_totals() -> dict:
    """Count / average VERPR / weighted cost for rows where Status='1'.

    Mirrors the Access form's AnzElem / DurchnKost / KostGew strip.
    """
    con = get_conn()
    try:
        row = con.execute(
            f'''SELECT COUNT(*) AS cnt,
                       AVG("VERPR") AS avg_verpr,
                       SUM("VERPR") AS sum_verpr
                FROM "{MAIN_TABLE}"
                WHERE CAST("Status" AS TEXT) = '1' '''
        ).fetchone()
    finally:
        con.close()
    cnt = row["cnt"] or 0
    avg = row["avg_verpr"] or 0.0
    return {
        "count": cnt,
        "avg_cost": round(avg, 4),
        "weighted_cost": round(avg * cnt, 2),
    }

"""Analysis queries ported from the Access application.

Column names use the sanitized SQLite identifiers produced by
migrate_from_access.py:
    SummevonZugang/Bedarf  -> SummevonZugang_Bedarf
    SummevonSummevonZugang/Bedarf (in *_einf_übersicht) -> SummevonSummevonZugang_Bedarf
"""
from __future__ import annotations

from .db import get_conn

TOP_N = 500


# --- cost ups / downs top 500 ----------------------------------------------
# Ports qry_cost_ups_and_dows_top_500_summe_zug_bed (aggregation) +
# qry_cost_ups_top_500_add / qry_cost_downs_top_500_add (final columns).

def _cost_rows(direction: str):
    """direction = 'up' or 'down'."""
    # The Access make-table first sums Zugang/Bedarf per (Material, Werk, VERPR,
    # VERPR_vm); then joins back to compute the cost delta * volume.
    if direction == "up":
        delta = '("VERPR" - "VERPR_vm")'
    else:
        delta = '("VERPR_vm" - "VERPR")'

    sql = f"""
        WITH agg AS (
            SELECT "Material", "Werk", "VERPR", "VERPR_vm",
                   SUM("SummevonZugang_Bedarf") AS sum_zb
            FROM tbl_alle_anzeigen
            GROUP BY "Material", "Werk", "VERPR", "VERPR_vm"
        ),
        named AS (
            SELECT a.*, aa."Materialkurztext", aa."Bestellmengeneinheit",
                   aa."ZPLP1", aa."ZPLP2", aa."ZPLP3", aa."GEWEI", aa."GROES",
                   aa."NTGEW", aa."MSTAE", aa."WRKST", aa."ZZ_SPEC_ID", aa."VPRSV"
            FROM agg a
            JOIN tbl_alle_anzeigen aa
              ON aa."Material" = a."Material" AND aa."Werk" = a."Werk"
             AND aa."VERPR" = a."VERPR" AND aa."VERPR_vm" = a."VERPR_vm"
            GROUP BY a."Material", a."Werk", a."VERPR", a."VERPR_vm"
        )
        SELECT "Material", "Materialkurztext", "Werk",
               sum_zb AS "SummevonSummevonZugang_Bedarf",
               "VERPR", "VERPR_vm",
               {delta} * sum_zb AS cost_value,
               CASE WHEN (sum_zb * "VERPR_vm") <> 0
                    THEN ({delta} * sum_zb) / (sum_zb * "VERPR_vm")
                    ELSE NULL END AS cost_pct,
               ("VERPR" - "VERPR_vm") AS cost_part,
               "Bestellmengeneinheit", "ZPLP1", "ZPLP2", "ZPLP3",
               "GEWEI", "GROES", "NTGEW", "MSTAE", "WRKST", "ZZ_SPEC_ID", "VPRSV"
        FROM named
        WHERE "VERPR_vm" <> 0
        ORDER BY {delta} * sum_zb DESC
        LIMIT {TOP_N}
    """
    con = get_conn()
    try:
        return [dict(r) for r in con.execute(sql).fetchall()]
    finally:
        con.close()


def cost_ups():
    return _cost_rows("up")


def cost_downs():
    return _cost_rows("down")


# --- plan price 1 / 2 -------------------------------------------------------
# qry_planpreis_1: Material NOT LIKE '11.5*'
# qry_planpreis_2: Material LIKE '11.5*'
# *_einf_übersicht: grouped overview with Sum(Zugang/Bedarf) and Sum(Summe).

def planpreis_overview(variant: int):
    op = "NOT LIKE" if variant == 1 else "LIKE"
    sql = f"""
        SELECT "Material", "Materialkurztext", "Werk", "Dispositionselement",
               "Bestellmengeneinheit",
               SUM("SummevonZugang_Bedarf") AS "SummevonSummevonZugang_Bedarf",
               "Lieferant", "VERPR", "VERPR_vm", "BWKEY",
               "ZPLP1", "ZPLP2", "ZPLP3", "GEWEI", "GROES", "NTGEW",
               "MSTAE", "WRKST", "ZZ_SPEC_ID", "ERSDA",
               SUM("Summe") AS "SummevonSumme",
               "Datentyp", "VPRSV", "STPRS", "Diff", "Status", "PEINH"
        FROM tbl_alle_anzeigen
        WHERE "Material" {op} '11.5%'
        GROUP BY "Material", "Materialkurztext", "Werk", "Dispositionselement",
                 "Bestellmengeneinheit", "Lieferant", "VERPR", "VERPR_vm", "BWKEY",
                 "ZPLP1", "ZPLP2", "ZPLP3", "GEWEI", "GROES", "NTGEW",
                 "MSTAE", "WRKST", "ZZ_SPEC_ID", "ERSDA",
                 "Datentyp", "VPRSV", "STPRS", "Diff", "Status", "PEINH"
        ORDER BY "Material"
    """
    con = get_conn()
    try:
        return [dict(r) for r in con.execute(sql).fetchall()]
    finally:
        con.close()
# --- demand -----------------------------------------------------------------
# qry_summe_demand: Direktes, schnelles Aggregieren aus der importierten Tabelle

def demand():
    sql = """
        SELECT "Material",
               SUM("SummevonSummevonZugang_Bedarf") AS "Demand"
        FROM "tbl_alle_anzeigen_einf_uebersicht"
        WHERE "Dispositionselement" IN ('BS-Ein', 'BS-Anf')
        GROUP BY "Material"
    """
    con = get_conn()
    try:
        return [dict(r) for r in con.execute(sql).fetchall()]
    finally:
        con.close()


# --- marked simple overview export -----------------------------------------
# qry_xls_ausgabe_markierte_einf_übersicht + demand join.
# Nutzt die native SQLite-Tabelle für maximale Performance.

def marked_overview_export():
    sql = """
        WITH dem AS (
            SELECT "Material",
                   SUM("SummevonSummevonZugang_Bedarf") AS demand
            FROM "tbl_alle_anzeigen_einf_uebersicht"
            WHERE "Dispositionselement" IN ('BS-Ein', 'BS-Anf')
            GROUP BY "Material"
        )
        SELECT e."Material", e."Materialkurztext",
               COALESCE(dem.demand, 0) AS "Demand",
               SUM(e."SummevonSummevonZugang_Bedarf") AS "Stock",
               e."VERPR", e."ZPLP1", e."ZPLP2", e."ZPLP3",
               e."Bestellmengeneinheit", e."Lieferant", e."BWKEY",
               e."GEWEI", e."GROES", e."NTGEW", e."MSTAE", e."WRKST",
               e."NORMT", e."ZZ_SPEC_ID", e."QZGTP", e."ERSDA",
               e."Datentyp", e."VPRSV", e."Status", e."VERBRL6M"
        FROM "tbl_alle_anzeigen_einf_uebersicht" e
        LEFT JOIN dem ON dem."Material" = e."Material"
        WHERE e."BWKEY" = '3100'
          AND CAST(e."Status" AS TEXT) = '1'
          AND e."Dispositionselement" = 'BStand'
        GROUP BY e."Material", e."Materialkurztext", e."VERPR",
                 e."ZPLP1", e."ZPLP2", e."ZPLP3", e."Bestellmengeneinheit",
                 e."Lieferant", e."BWKEY", e."GEWEI", e."GROES", e."NTGEW",
                 e."MSTAE", e."WRKST", e."NORMT", e."ZZ_SPEC_ID", e."QZGTP",
                 e."ERSDA", e."Datentyp", e."VPRSV", e."Status", e."VERBRL6M"
        ORDER BY e."Material"
    """
    con = get_conn()
    try:
        return [dict(r) for r in con.execute(sql).fetchall()]
    finally:
        con.close()
# --- new / locked parts -----------------------------------------------------
# qry_xls_export_newparts_parts (Datentyp='New') /
# qry_xls_export_lockedparts_parts (Datentyp='locked').

def parts_by_datentyp(datentyp: str):
    sql = """
        SELECT "Material", "Materialkurztext", "Werk", "Dispositionselement",
               "Bestellmengeneinheit",
               "SummevonZugang_Bedarf", "Lieferant", "VERPR", "VERPR_vm",
               "BWKEY", "ZPLP1", "ZPLP2", "ZPLP3", "GEWEI", "GROES", "NTGEW",
               "MSTAE", "WRKST", "ZZ_SPEC_ID", "ERSDA", "Summe", "VPRSV",
               "STPRS", "Diff", "Status", "PEINH", "VERBRL6M", "NORMT"
        FROM tbl_alle_anzeigen
        WHERE "Datentyp" = ?
        GROUP BY "Material", "Materialkurztext", "Werk", "Dispositionselement",
                 "Bestellmengeneinheit", "SummevonZugang_Bedarf", "Lieferant",
                 "VERPR", "VERPR_vm", "BWKEY", "ZPLP1", "ZPLP2", "ZPLP3",
                 "GEWEI", "GROES", "NTGEW", "MSTAE", "WRKST", "ZZ_SPEC_ID",
                 "ERSDA", "Summe", "VPRSV", "STPRS", "Diff", "Status",
                 "PEINH", "VERBRL6M", "NORMT"
        ORDER BY "Material"
    """
    con = get_conn()
    try:
        return [dict(r) for r in con.execute(sql, (datentyp,)).fetchall()]
    finally:
        con.close()

"""FastAPI web replacement for the Access Product Cost Management tool."""
from __future__ import annotations

import os
import hashlib

from fastapi import FastAPI, Form, Request, Depends, HTTPException, status
from fastapi.responses import HTMLResponse, JSONResponse, Response, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import analysis, db, excel

# --- Pfade definieren ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)

# --- FastAPI App initialisieren ---
app = FastAPI(title="Product Cost Management")

# --- Statische Dateien und Templates einbinden ---
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(PARENT_DIR, "templates"))

# FIX FÜR PYTHON 3.14: Verhindert den Cache-Absturz von Jinja2
templates.env.cache = None

# --- Passwort-Schutz (Direkter Text-Abgleich für lokalen Test) ---
USER_DB = {
    "admin": "geheim123"
}

def verify_password(plain_password: str, stored_password: str) -> bool:
    return plain_password.strip() == stored_password.strip()

def get_current_user(request: Request):
    """Sicherheits-Check: Prüft ob ein gültiges Session-Cookie existiert."""
    username = request.cookies.get("session_user")
    if not username or username not in USER_DB:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Nicht angemeldet"
        )
    return username

def _xlsx(data: bytes, filename: str) -> Response:
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

# --- Authentifizierungs-Routen ---

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    # Wenn nicht eingeloggt, zur Anmeldung schicken
    if request.cookies.get("session_user") not in USER_DB:
        return templates.TemplateResponse(request=request, name="login.html")
        
    # --- LIVE DATEN AUS IHRER SQLITE BERECHNEN ---
    ups_data = analysis.cost_ups()  # Holt die echten Top-Kostensteigerungen
    downs_data = analysis.cost_downs()  # Holt die echten Top-Kostensenkungen
    
    # Echte Kennzahlen berechnen
    echte_savings = sum(r['cost_value'] for r in downs_data if r['cost_value'] is not None)
    echte_erhoehungen = sum(r['cost_value'] for r in ups_data if r['cost_value'] is not None)
    
    # Top 5 Listen für das Dashboard zuschneiden
    top_ups = ups_data[:5]
    top_downs = downs_data[:5]
    
    # Daten an das HTML-Dashboard übergeben
    return templates.TemplateResponse(
        request=request, 
        name="home.html",
        context={
            "savings": f"${round(echte_savings, 2):,}",
            "losses": f"${round(echte_erhoehungen, 2):,}",
            "top_ups": top_ups,
            "top_downs": top_downs
        }
    )


@app.post("/login")
def login(request: Request, username: str = Form(...), password: str = Form(...)):
    if username in USER_DB and verify_password(password, USER_DB[username]):
        response = RedirectResponse(url="/data", status_code=status.HTTP_303_SEE_OTHER)
        response.set_cookie(key="session_user", value=username, httponly=True)
        return response
    
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"error": "Ungültiger Benutzername oder Passwort."}
    )

@app.get("/logout")
def logout():
    response = RedirectResponse(url="/")
    response.delete_cookie("session_user")
    return response
# --- main data grid (Alle Routen durch Login geschützt) -------------------

@app.get("/data", response_class=HTMLResponse)
def data_grid(
    request: Request,
    material: str = "",
    werk: str = "",
    lieferant: str = "",
    datentyp: str = "",
    status: str = "",
    sort: str = "",
    direction: str = "asc",
    page: int = 1,
    user: str = Depends(get_current_user),
):
    page = max(1, page)
    per_page = 100
    filters = {
        "Material": material, "Werk": werk,
        "Lieferant": lieferant, "Datentyp": datentyp,
    }
    rows, total = db.fetch_rows(
        filters=filters, status=status or None,
        sort=sort or None, direction=direction,
        limit=per_page, offset=(page - 1) * per_page,
    )
    pages = max(1, (total + per_page - 1) // per_page)
    
    return templates.TemplateResponse(
        request=request,
        name="data.html",
        context={
            "rows": rows,
            "labels": db.column_labels(),
            "grid_columns": db.GRID_COLUMNS,
            "totals": db.marked_totals(),
            "total": total,
            "page": page,
            "pages": pages,
            "filters": {"material": material, "werk": werk,
                        "lieferant": lieferant, "datentyp": datentyp,
                        "status": status},
            "sort": sort,
            "direction": direction,
        },
    )


@app.get("/data/{rowid}/detail", response_class=HTMLResponse)
def row_detail(request: Request, rowid: int, user: str = Depends(get_current_user)):
    row = db.fetch_one(rowid)
    if row is None:
        return HTMLResponse("<p>Row not found.</p>", status_code=404)
        
    return templates.TemplateResponse(
        request=request,
        name="detail.html",
        context={"row": dict(row), "labels": db.column_labels()}
    )


@app.post("/data/{rowid}/status")
def toggle_status(rowid: int, value: int = Form(...), user: str = Depends(get_current_user)):
    db.set_status(rowid, 1 if value else 0)
    return JSONResponse({"rowid": rowid, "status": 1 if value else 0,
                         "totals": db.marked_totals()})


@app.post("/data/status/reset")
def reset_status(user: str = Depends(get_current_user)):
    n = db.reset_all_status()
    return JSONResponse({"reset": n, "totals": db.marked_totals()})
# --- analysis views (Durch Login geschützt) -------------------------------

@app.get("/analysis/cost", response_class=HTMLResponse)
def cost_analysis(request: Request, user: str = Depends(get_current_user)):
    return templates.TemplateResponse(
        request=request,
        name="cost.html",
        context={"ups": analysis.cost_ups(), "downs": analysis.cost_downs()}
    )


@app.get("/planpreis/{variant}", response_class=HTMLResponse)
def planpreis(request: Request, variant: int, user: str = Depends(get_current_user)):
    if variant not in (1, 2):
        return HTMLResponse("Unknown variant", status_code=404)
    rows = analysis.planpreis_overview(variant)
    
    return templates.TemplateResponse(
        request=request,
        name="planpreis.html",
        context={"variant": variant, "rows": rows, "labels": db.column_labels()}
    )


# --- exports (Durch Login geschützt) ---------------------------------------

@app.get("/export/cost.xlsx")
def export_cost(user: str = Depends(get_current_user)):
    data = excel.cost_report(analysis.cost_ups(), analysis.cost_downs())
    return _xlsx(data, "Report_cost_ups_and_downs.xlsx")


@app.get("/export/marked.xlsx")
def export_marked(user: str = Depends(get_current_user)):
    rows = analysis.marked_overview_export()
    data = excel.generic_report(
        "markierte_einf_uebersicht", rows,
        currency_cols=("VERPR", "ZPLP1", "ZPLP2", "ZPLP3", "Stock"),
    )
    return _xlsx(data, "Report_marked_simple_overview.xlsx")


@app.get("/export/parts.xlsx")
def export_parts(user: str = Depends(get_current_user)):
    new = analysis.parts_by_datentyp("New")
    locked = analysis.parts_by_datentyp("locked")
    data = excel.two_sheet_report(
        "new_parts", new, "locked_parts", locked,
        currency_cols=("VERPR", "VERPR_vm", "STPRS", "Diff"),
    )
    return _xlsx(data, "Report_newparts_lockedparts.xlsx")


@app.get("/export/demand.xlsx")
def export_demand(user: str = Depends(get_current_user)):
    data = excel.generic_report("demand", analysis.demand())
    return _xlsx(data, "Report_demand.xlsx")

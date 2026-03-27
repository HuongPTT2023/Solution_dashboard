from datetime import datetime
from io import BytesIO, StringIO
import os
import secrets
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from fastapi import Depends, FastAPI, HTTPException, Request, status
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

app = FastAPI()
security = HTTPBasic()
VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")

SHEET_ID = "1SSLaOM1YQ7xQ6E2wVrRiZ8BBYjNsbGMpvDdhpcZk5pU"

DASHBOARD_USERNAME = os.getenv("DASHBOARD_USERNAME", "admin")
DASHBOARD_PASSWORD = os.getenv("DASHBOARD_PASSWORD", "change-me")

SHEETS = {
    "config": "1346054572",
    "kpis": "0",
    "projects": "1056309973",
    "platforms": "748315146",
    "tests": "1077608661",
}

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


def require_auth(credentials: HTTPBasicCredentials = Depends(security)):
    username_ok = secrets.compare_digest(credentials.username, DASHBOARD_USERNAME)
    password_ok = secrets.compare_digest(credentials.password, DASHBOARD_PASSWORD)
    if not (username_ok and password_ok):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Unauthorized",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


def load_sheet_df(gid: str):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    response.encoding = "utf-8"
    return pd.read_csv(StringIO(response.text)).fillna("")


def load_sheet_csv(gid: str):
    return load_sheet_df(gid).to_dict(orient="records")


def normalize_text(value):
    return str(value).strip() if value is not None else ""


def normalize_int(value, default=0):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def normalize_project(row):
    return {
        "name": normalize_text(row.get("Name")),
        "product_line": normalize_text(row.get("Product_line")),
        "model": normalize_text(row.get("Model")),
        "spec": normalize_text(row.get("Spec")),
        "chipset": normalize_text(row.get("Chipset")),
        "folder": normalize_text(row.get("Folder")),
        "operation_model": normalize_text(row.get("Operation_model")),
        "current_stage": normalize_text(row.get("Current_stage")),
        "concept_quarter": normalize_text(row.get("Concept")),
        "research_quarter": normalize_text(row.get("Research")),
        "prototype_quarter": normalize_text(row.get("Prototype")),
        "system_test_quarter": normalize_text(row.get("System_test")),
        "poc_quarter": normalize_text(row.get("POC")),
        "mp_quarter": normalize_text(row.get("MP")),
        "end_quarter": normalize_text(row.get("End")),
        "progress_pct": normalize_int(row.get("Progress", 0)),
        "status": normalize_text(row.get("Status")),
        "risk_note": normalize_text(row.get("Risk_note")),
        "bom_cost": normalize_text(
            row.get("BomCost") or row.get("Bom_cost") or row.get("BOMCost") or row.get("BOM_Cost")
        ),
        "objective": normalize_text(row.get("Objective")),
        "result_summary": normalize_text(row.get("Result_summary")),
        "dependency": normalize_text(row.get("Dependency")),
        "action_plan": normalize_text(row.get("Action_plan")),
        "highlight_tag": normalize_text(row.get("Highlight_tag")),
        "image_url": normalize_text(row.get("image_url")),
    }


def normalize_status(value):
    text = normalize_text(value).lower()
    if text == "on track":
        return "On Track"
    if text == "at risk":
        return "At Risk"
    if text in {"pending", "blocked"}:
        return "Pending"
    if text == "stop":
        return "Stop"
    return ""


def build_config(rows):
    config = {}
    for row in rows:
        key = normalize_text(row.get("key"))
        value = normalize_text(row.get("value"))
        if key:
            config[key] = value
    return config


def build_kpis(projects):
    statuses = [normalize_status(project.get("status")) for project in projects]
    return [
        {
            "metric": "Du_an_dang_nghien_cuu",
            "value": len(projects),
            "icon": "project",
        },
        {
            "metric": "On_Track",
            "value": sum(1 for status in statuses if status == "On Track"),
            "icon": "check",
        },
        {
            "metric": "At_Risk",
            "value": sum(1 for status in statuses if status == "At Risk"),
            "icon": "warning",
        },
        {
            "metric": "Pending",
            "value": sum(1 for status in statuses if status == "Pending"),
            "icon": "pending",
        },
        {
            "metric": "Stop",
            "value": sum(1 for status in statuses if status == "Stop"),
            "icon": "stop",
        },
    ]


def get_dashboard_data():
    config_rows = load_sheet_csv(SHEETS["config"])
    raw_projects = load_sheet_csv(SHEETS["projects"])

    projects = [normalize_project(row) for row in raw_projects]
    kpis = build_kpis(projects)
    refreshed_at = datetime.now(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")

    return {
        "config": build_config(config_rows),
        "kpis": kpis,
        "projects": projects,
        "platforms": load_sheet_csv(SHEETS["platforms"]),
        "tests": load_sheet_csv(SHEETS["tests"]),
        "refreshed_at": refreshed_at,
    }


def build_export_workbook():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df = load_sheet_df(SHEETS["projects"])
        df.to_excel(writer, sheet_name="Projects", index=False)
    output.seek(0)
    return output


@app.get("/")
def read_root(username: str = Depends(require_auth)):
    return {"message": f"Dashboard backend is running for {username}"}


@app.get("/api/data")
def api_data(username: str = Depends(require_auth)):
    return get_dashboard_data()


@app.get("/api/export/google-sheet.xlsx")
def export_google_sheet(username: str = Depends(require_auth)):
    workbook = build_export_workbook()
    filename = f"solution-dashboard-full-data-{datetime.now(VN_TZ).strftime('%Y%m%d-%H%M%S')}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return StreamingResponse(
        workbook,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request, username: str = Depends(require_auth)):
    return templates.TemplateResponse(
        "dashboard.html",
        {
            "request": request,
            "username": username,
        },
    )

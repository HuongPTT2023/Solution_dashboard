from io import StringIO

import pandas as pd
import requests
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

app = FastAPI()

SHEET_ID = "1SSLaOM1YQ7xQ6E2wVrRiZ8BBYjNsbGMpvDdhpcZk5pU"

SHEETS = {
    "config": "1346054572",
    "kpis": "0",
    "projects": "1056309973",
    "platforms": "748315146",
    "tests": "1077608661",
}

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


def load_sheet_csv(gid: str):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    response.encoding = "utf-8"
    df = pd.read_csv(StringIO(response.text))
    df = df.fillna("")
    return df.to_dict(orient="records")


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
        "bom_cost": normalize_text(row.get("BomCost") or row.get("Bom_cost") or row.get("BOMCost") or row.get("BOM_Cost")),
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
    if text == "blocked":
        return "Blocked"
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
            "metric": "Blocked",
            "value": sum(1 for status in statuses if status == "Blocked"),
            "icon": "block",
        },
    ]


def get_dashboard_data():
    config_rows = load_sheet_csv(SHEETS["config"])
    raw_projects = load_sheet_csv(SHEETS["projects"])

    projects = [normalize_project(row) for row in raw_projects]
    kpis = build_kpis(projects)

    return {
        "config": build_config(config_rows),
        "kpis": kpis,
        "projects": projects,
        "platforms": load_sheet_csv(SHEETS["platforms"]),
        "tests": load_sheet_csv(SHEETS["tests"]),
    }


@app.get("/")
def read_root():
    return {"message": "Dashboard backend is running"}


@app.get("/api/data")
def api_data():
    return get_dashboard_data()


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    return templates.TemplateResponse(
        "dashboard.html",
        {
            "request": request,
        },
    )

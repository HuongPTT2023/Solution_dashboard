from datetime import datetime
from io import BytesIO, StringIO
import json
import os
import secrets
import re
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from zoneinfo import ZoneInfo

from fastapi import Depends, FastAPI, HTTPException, Request, status
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
import gspread
import pandas as pd
import requests

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

PROJECT_UPDATE_FIELDS = {
    "risk_note": "Risk_note",
    "objective": "Objective",
    "operation_model": "Operation_model",
    "spec": "Spec",
    "chipset": "Chipset",
    "result_summary": "Result_summary",
    "dependency": "Dependency",
    "action_plan": "Action_plan",
    "folder": "Folder",
}

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


class ProjectUpdatePayload(BaseModel):
    project_id: str
    name: str
    model: str
    risk_note: str = ""
    objective: str = ""
    operation_model: str = ""
    spec: str = ""
    chipset: str = ""
    result_summary: str = ""
    dependency: str = ""
    action_plan: str = ""
    folder: str = ""


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


def get_gspread_client():
    service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()

    if service_account_json:
        return gspread.service_account_from_dict(json.loads(service_account_json))
    if service_account_file:
        return gspread.service_account(filename=service_account_file)

    raise HTTPException(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        detail=(
            "Chua cau hinh Google Sheets write-back. "
            "Can set GOOGLE_SERVICE_ACCOUNT_JSON hoac GOOGLE_SERVICE_ACCOUNT_FILE."
        ),
    )


def get_projects_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open_by_key(SHEET_ID)
    return spreadsheet.get_worksheet_by_id(int(SHEETS["projects"]))


def normalize_text(value):
    return str(value).strip() if value is not None else ""


def normalize_int(value, default=0):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def normalize_image_url(value):
    image_url = normalize_text(value)
    if not image_url:
        return ""

    parsed = urlparse(image_url)
    if not parsed.scheme or not parsed.netloc:
        return image_url

    host = parsed.netloc.lower()
    if "1drv.ms" in host or "onedrive.live.com" in host or "sharepoint.com" in host:
        query = dict(parse_qsl(parsed.query, keep_blank_values=True))
        query["download"] = "1"
        return urlunparse(parsed._replace(query=urlencode(query)))

    if "drive.google.com" in host:
        if parsed.path.startswith("/thumbnail") or parsed.path.startswith("/uc"):
            return image_url

        match = re.search(r"/file/d/([^/]+)", parsed.path)
        if match:
            file_id = match.group(1)
            return f"https://drive.google.com/thumbnail?id={file_id}&sz=w1000"

        query = dict(parse_qsl(parsed.query, keep_blank_values=True))
        file_id = query.get("id", "")
        if file_id:
            return f"https://drive.google.com/thumbnail?id={file_id}&sz=w1000"

    return image_url


def normalize_project(row):
    image_url = normalize_image_url(row.get("image_url") or row.get("Image_URL") or row.get("Image_url"))
    return {
        "project_id": normalize_text(row.get("Project_ID")),
        "name": normalize_text(row.get("Name")),
        "product_line": normalize_text(row.get("Product_line")),
        "wifi_standard": normalize_text(
            row.get("Wi-Fi_Standard") or row.get("WiFi_Standard") or row.get("Wifi_Standard")
        ),
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
        "last_updated": normalize_text(row.get("Last_Updated")),
        "highlight_tag": normalize_text(row.get("Highlight_tag")),
        "image_url": image_url,
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


def update_project_in_sheet(payload: ProjectUpdatePayload):
    worksheet = get_projects_worksheet()
    records = worksheet.get_all_records()

    target_index = None
    for idx, row in enumerate(records, start=2):
        project_id_matches = normalize_text(row.get("Project_ID")) == normalize_text(payload.project_id)
        name_matches = normalize_text(row.get("Name")) == normalize_text(payload.name)
        model_matches = normalize_text(row.get("Model")) == normalize_text(payload.model)
        if project_id_matches and name_matches and model_matches:
            target_index = idx
            break

    if target_index is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Khong tim thay dong du an theo Project_ID + Name + Model trong sheet Projects.",
        )

    header = worksheet.row_values(1)
    updates = []
    payload_data = payload.model_dump()

    for field, column_name in PROJECT_UPDATE_FIELDS.items():
        if column_name not in header:
            continue
        col_index = header.index(column_name) + 1
        updates.append(
            {
                "range": gspread.utils.rowcol_to_a1(target_index, col_index),
                "values": [[normalize_text(payload_data.get(field, ""))]],
            }
        )

    if "Last_Updated" in header:
        last_updated_col = header.index("Last_Updated") + 1
        updates.append(
            {
                "range": gspread.utils.rowcol_to_a1(target_index, last_updated_col),
                "values": [[datetime.now(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")]],
            }
        )

    if updates:
        worksheet.batch_update(updates, value_input_option="USER_ENTERED")


@app.get("/")
def read_root(username: str = Depends(require_auth)):
    return {"message": f"Dashboard backend is running for {username}"}


@app.get("/api/data")
def api_data(username: str = Depends(require_auth)):
    return get_dashboard_data()


@app.post("/api/project/update")
def update_project(payload: ProjectUpdatePayload, username: str = Depends(require_auth)):
    update_project_in_sheet(payload)
    data = get_dashboard_data()
    return {
        "message": f"Updated by {username}",
        "project": next(
            (
                project
                for project in data["projects"]
                if normalize_text(project.get("project_id")) == normalize_text(payload.project_id)
                and normalize_text(project.get("name")) == normalize_text(payload.name)
                and normalize_text(project.get("model")) == normalize_text(payload.model)
            ),
            None,
        ),
        "refreshed_at": data["refreshed_at"],
    }


@app.get("/api/export/google-sheet.xlsx")
def export_google_sheet(username: str = Depends(require_auth)):
    workbook = build_export_workbook()
    filename = f"solution-dashboard-projects-{datetime.now(VN_TZ).strftime('%Y%m%d-%H%M%S')}.xlsx"
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

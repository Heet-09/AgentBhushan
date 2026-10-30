"""Run the local Jira bug review interface."""

import json
import os
import uuid
from pathlib import Path

import pandas as pd
import requests
from flask import Flask, render_template, request

from create_jira_bugs import build_issue, load_env_file, record_id
from parse_excel import get_failed_rows


APP_DIR = Path(__file__).resolve().parent
CREATED_FILE = APP_DIR / "created_bugs.json"
MAX_UPLOAD_SIZE = 16 * 1024 * 1024
ANALYSES = {}

load_env_file(str(APP_DIR / ".env"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE


def page(**context):
    context.setdefault("project", os.environ.get("JIRA_PROJECT", ""))
    context.setdefault("jira_configured", bool(os.environ.get("JIRA_URL") and os.environ.get("JIRA_API_TOKEN")))
    return render_template("index.html", **context)


@app.get("/")
def index():
    return page()


@app.post("/analyze")
def analyze():
    report = request.files.get("report")
    if not report or not report.filename:
        return page(error="Choose an Excel report to analyze."), 400
    if Path(report.filename).suffix.lower() not in {".xlsx", ".xlsm"}:
        return page(error="Upload an .xlsx or .xlsm Excel file."), 400

    try:
        sheets = pd.read_excel(report.stream, sheet_name=None, header=None)
        failures = get_failed_rows(sheets)
    except Exception as error:
        return page(error=f"Could not read this workbook: {error}"), 400

    run_id = uuid.uuid4().hex
    ANALYSES[run_id] = {"filename": Path(report.filename).name, "records": failures}
    while len(ANALYSES) > 20:
        ANALYSES.pop(next(iter(ANALYSES)))

    return page(
        run_id=run_id,
        filename=Path(report.filename).name,
        records=failures,
        selected_indices={str(index) for index in range(len(failures))},
    )


@app.post("/create")
def create():
    run_id = request.form.get("run_id", "")
    analysis = ANALYSES.get(run_id)
    if analysis is None:
        return page(error="This review has expired. Upload the report again."), 400

    records = analysis["records"]
    selected_indices = {
        int(value)
        for value in request.form.getlist("rows")
        if value.isdigit() and int(value) < len(records)
    }
    project = request.form.get("project", "").strip()
    if not project:
        return page(
            run_id=run_id,
            filename=analysis["filename"],
            records=records,
            selected_indices={str(index) for index in selected_indices},
            error="Enter a Jira project key before creating bugs.",
        ), 400
    if not selected_indices:
        return page(
            run_id=run_id,
            filename=analysis["filename"],
            records=records,
            selected_indices=set(),
            error="Select at least one failed test case.",
        ), 400

    base_url = os.environ.get("JIRA_URL", "").rstrip("/")
    token = os.environ.get("JIRA_API_TOKEN")
    email = os.environ.get("JIRA_EMAIL")
    if not base_url or not token:
        return page(
            run_id=run_id,
            filename=analysis["filename"],
            records=records,
            selected_indices={str(index) for index in selected_indices},
            project=project,
            error="Set JIRA_URL and JIRA_API_TOKEN in the .env file before creating bugs.",
        ), 400

    created = {}
    if CREATED_FILE.exists():
        try:
            created = json.loads(CREATED_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return page(
                run_id=run_id,
                filename=analysis["filename"],
                records=records,
                selected_indices={str(index) for index in selected_indices},
                project=project,
                error="Could not read created_bugs.json. Fix or remove that file before continuing.",
            ), 500

    session = requests.Session()
    if email:
        session.auth = (email, token)
    else:
        session.headers["Authorization"] = f"Bearer {token}"

    results = []
    for index in sorted(selected_indices):
        record = records[index]
        source_id = record_id(record)
        if source_id in created:
            results.append({"state": "skipped", "message": f"{source_id} was already created as {created[source_id]}"})
            continue

        try:
            response = session.post(
                f"{base_url}/rest/api/2/issue",
                json=build_issue(record, project),
                timeout=30,
            )
            if response.ok:
                issue_key = response.json()["key"]
                created[source_id] = issue_key
                CREATED_FILE.write_text(json.dumps(created, indent=2), encoding="utf-8")
                results.append({"state": "created", "message": f"{source_id} created as {issue_key}"})
            else:
                details = response.text.replace("\n", " ")[:300]
                results.append({"state": "failed", "message": f"{source_id}: Jira returned {response.status_code}: {details}"})
        except (requests.RequestException, ValueError, KeyError) as error:
            results.append({"state": "failed", "message": f"{source_id}: {error}"})

    return page(
        run_id=run_id,
        filename=analysis["filename"],
        records=records,
        selected_indices={str(index) for index in selected_indices},
        project=project,
        results=results,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
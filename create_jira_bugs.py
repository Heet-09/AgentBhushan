"""Create one Jira Bug for every failed test case in failed.json.

Field mapping:
    summary     <- "Test Case"
    description <- "Steps" + "Expected Result"
    issuetype   <- Bug
    priority    <- "Priority"

By default this is a DRY RUN: it only prints what would be created.
Pass --create to really create the bugs.

Configuration (environment variables, or the same KEY=value lines in a .env file next to this script):
    JIRA_URL        e.g. https://yourcompany.atlassian.net
    JIRA_PROJECT    project key, e.g. SCAD (or use --project)
    JIRA_EMAIL      Jira Cloud account email (omit for a Server/Data Center personal access token)
    JIRA_API_TOKEN  API token (Cloud) or personal access token (Server/Data Center)

Usage:
    python create_jira_bugs.py failed.json                  # dry run
    python create_jira_bugs.py failed.json --create         # create the bugs
"""

import argparse
import json
import os
import sys

import requests

CREATED_FILE = "created_bugs.json"  # remembers what was created so a re-run does not duplicate


def load_env_file(path):
    """Load KEY=value lines from a .env file into os.environ (real environment variables win)."""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def build_issue(record, project):
    description = f"*Steps:*\n{record.get('Steps', '')}\n\n*Expected Result:*\n{record.get('Expected Result', '')}"
    fields = {
        "project": {"key": project},
        "summary": record["Test Case"][:255],  # Jira's summary limit
        "description": description,
        "issuetype": {"name": "Bug"},
    }
    if record.get("Priority"):
        fields["priority"] = {"name": record["Priority"]}
    return {"fields": fields}


def record_id(record):
    return f"{record.get('Sheet')}!{record.get('Excel row')}"


def main():
    load_env_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
    parser = argparse.ArgumentParser(description="Create Jira bugs from failed test cases.")
    parser.add_argument("path", help="JSON file made by parse_excel.py --failed")
    parser.add_argument("--project", default=os.environ.get("JIRA_PROJECT"), help="Jira project key")
    parser.add_argument("--create", action="store_true", help="Really create the bugs (default is a dry run)")
    args = parser.parse_args()

    with open(args.path, encoding="utf-8") as f:
        records = json.load(f)

    created = {}
    if os.path.exists(CREATED_FILE):
        with open(CREATED_FILE, encoding="utf-8") as f:
            created = json.load(f)

    if not args.project:
        sys.exit("Set the project key with --project or JIRA_PROJECT.")

    base_url = os.environ.get("JIRA_URL", "").rstrip("/")
    token = os.environ.get("JIRA_API_TOKEN")
    email = os.environ.get("JIRA_EMAIL")
    if args.create and not (base_url and token):
        sys.exit("Set JIRA_URL and JIRA_API_TOKEN (and JIRA_EMAIL for Jira Cloud) before using --create.")

    session = requests.Session()
    if email:
        session.auth = (email, token)
    elif token:
        session.headers["Authorization"] = f"Bearer {token}"

    for record in records:
        key = record_id(record)
        if key in created:
            print(f"SKIP  {key}: already created as {created[key]}")
            continue

        issue = build_issue(record, args.project)
        if not args.create:
            print(f"DRY RUN {key}")
            print(json.dumps(issue["fields"], indent=2, ensure_ascii=False), "\n")
            continue

        response = session.post(f"{base_url}/rest/api/2/issue", json=issue)
        if response.ok:
            created[key] = response.json()["key"]
            with open(CREATED_FILE, "w", encoding="utf-8") as f:
                json.dump(created, f, indent=2)
            print(f"CREATED {key} -> {created[key]}")
        else:
            print(f"FAILED  {key}: {response.status_code} {response.text}")

    if not args.create:
        print("Dry run only. Nothing was created. Re-run with --create to create these bugs.")


if __name__ == "__main__":
    main()

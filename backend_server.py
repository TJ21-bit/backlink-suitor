"""
╔══════════════════════════════════════════════════════════╗
║         BACKLINK AUDITOR — Flask Backend Server v4       ║
╚══════════════════════════════════════════════════════════╝

RUN LOCALLY:
    python backend_server.py          → http://localhost:8080

RUN ON RENDER:
    Deployed with the Dockerfile in this folder (see DEPLOY.md).

SETTINGS (environment variables, or a local .env file — never commit it):
    TELEGRAM_TOKEN   Telegram bot token
    EMAIL_SENDER     Gmail address that sends reports
    EMAIL_PASSWORD   Gmail App Password
    APP_PASSWORD     If set, the site asks for this password (username: anything)
    BROWSER_WORKERS  Parallel headless Chrome tabs (default 3; use 2 on Render Standard)
"""

import os
import uuid
import hmac
import threading
import requests as req_lib
from flask import Flask, request, jsonify, send_file, Response

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_dotenv(path):
    """Read KEY=VALUE lines from a local .env file (no extra package needed)."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))
    except FileNotFoundError:
        pass


_load_dotenv(os.path.join(BASE_DIR, ".env"))

# ── Telegram Config ────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")

# Default recipients (always notified)
DEFAULT_CHAT_IDS = {
    "6595431574": "Tijeet (Team Lead)",
    "702339188" : "Nitish (Manager)",
}

def send_telegram(message: str, chat_ids: dict = None):
    """Send message to specified chat_ids or DEFAULT_CHAT_IDS."""
    if not TELEGRAM_TOKEN:
        print("Telegram skipped — TELEGRAM_TOKEN not set")
        return
    targets = chat_ids if chat_ids else DEFAULT_CHAT_IDS
    url     = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    for chat_id in targets:
        try:
            req_lib.post(url, json={
                "chat_id"   : chat_id,
                "text"      : message,
                "parse_mode": "HTML",
            }, timeout=10)
        except Exception as e:
            print(f"Telegram error for {chat_id}: {e}")


# ── Email Config ───────────────────────────────────────────────────────────
# Gmail SMTP — set EMAIL_SENDER and EMAIL_PASSWORD (a Gmail App Password)
# Get App Password: Google Account → Security → 2-Step Verification → App passwords
EMAIL_SENDER   = os.environ.get("EMAIL_SENDER", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")

def send_email(to_email: str, subject: str, body: str):
    """Send email via Gmail SMTP."""
    if not EMAIL_PASSWORD:
        print(f"Email skipped — EMAIL_PASSWORD not set (to: {to_email})")
        return
    try:
        import smtplib
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart

        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = EMAIL_SENDER
        msg["To"]      = to_email

        # Plain text version
        text_part = MIMEText(body, "plain")
        # HTML version
        html_body = body.replace("\n", "<br>").replace("<b>", "<strong>").replace("</b>", "</strong>")
        html_part = MIMEText(f"<html><body style='font-family:Arial;font-size:14px;'>{html_body}</body></html>", "html")

        msg.attach(text_part)
        msg.attach(html_part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, to_email, msg.as_string())
        print(f"✅ Email sent to {to_email}")
    except Exception as e:
        print(f"❌ Email error to {to_email}: {e}")


def build_telegram_summary(job_id: str, jobs: dict, filename: str) -> dict:
    """
    Scan the audited report and build detailed summaries per role.
    Returns { manager_msg, teamlead_msg, employee_msgs: {employee: msg} }
    """
    import pandas as pd, os

    job         = jobs[job_id]
    output_path = job.get("output", "")
    good        = job["good"]
    flagged     = job["flagged"]
    warn        = job["warn"]
    total       = good + flagged + warn

    # Try to read the actual output file for detailed analysis
    employee_data   = {}   # { employee: {good, flagged, warn, issues} }
    manager_data    = {}   # { manager: {good, flagged, warn} }
    teamlead_data   = {}   # { teamlead: {good, flagged, warn} }
    flag_reasons    = {}   # { reason: count }
    project_summary = {}   # { website/project: {good, flagged} }

    try:
        if output_path and os.path.exists(output_path):
            xl  = pd.ExcelFile(output_path)
            df  = pd.read_excel(output_path, sheet_name=xl.sheet_names[0], dtype=str).fillna("")

            for _, row in df.iterrows():
                result   = str(row.get("Audit Result",    ""))
                emp      = str(row.get("Employee Name",   "Unknown")).strip()
                mgr      = str(row.get("Project Manager", "")).strip()
                tl       = str(row.get("Team Lead",       "")).strip()
                website  = str(row.get("Website",         "")).strip()
                reason   = str(row.get("Flag Reason",     "")).strip()
                url      = str(row.get("Backlink Live Link", "")).strip()
                keyword  = str(row.get("Target Keyword",  "")).strip()
                activity = str(row.get("Backlink Activity Type", "")).strip()

                is_good    = "Good"    in result
                is_flagged = "Flagged" in result
                is_warn    = "Manual"  in result

                # Per employee
                if emp not in employee_data:
                    employee_data[emp] = {"good": 0, "flagged": 0, "warn": 0, "issues": []}
                if is_good:    employee_data[emp]["good"]    += 1
                if is_flagged: employee_data[emp]["flagged"] += 1
                if is_warn:    employee_data[emp]["warn"]    += 1
                if is_flagged and reason and reason != "All checks passed ✓":
                    employee_data[emp]["issues"].append({
                        "url"     : url[:60],
                        "keyword" : keyword,
                        "reason"  : reason[:80],
                        "activity": activity,
                    })

                # Per manager
                if mgr:
                    if mgr not in manager_data:
                        manager_data[mgr] = {"good": 0, "flagged": 0, "warn": 0}
                    if is_good:    manager_data[mgr]["good"]    += 1
                    if is_flagged: manager_data[mgr]["flagged"] += 1
                    if is_warn:    manager_data[mgr]["warn"]    += 1

                # Per team lead
                if tl:
                    if tl not in teamlead_data:
                        teamlead_data[tl] = {"good": 0, "flagged": 0, "warn": 0}
                    if is_good:    teamlead_data[tl]["good"]    += 1
                    if is_flagged: teamlead_data[tl]["flagged"] += 1
                    if is_warn:    teamlead_data[tl]["warn"]    += 1

                # Per project
                if website:
                    if website not in project_summary:
                        project_summary[website] = {"good": 0, "flagged": 0, "warn": 0}
                    if is_good:    project_summary[website]["good"]    += 1
                    if is_flagged: project_summary[website]["flagged"] += 1
                    if is_warn:    project_summary[website]["warn"]    += 1

                # Flag reasons
                if is_flagged and reason and reason != "All checks passed ✓":
                    for r in reason.split(";"):
                        r = r.strip()
                        if r:
                            flag_reasons[r] = flag_reasons.get(r, 0) + 1

    except Exception as e:
        print(f"Report scan error: {e}")

    # ── Build manager message (full overview) ─────────────────────────────
    pct_good = round((good / total * 100)) if total else 0

    manager_msg = (
        f"📊 <b>Backlink Audit Report</b>\n"
        f"{'─' * 32}\n"
        f"🆔 Job: <code>{job_id}</code>\n"
        f"📁 File: {filename}\n\n"
        f"<b>Overall Summary:</b>\n"
        f"📌 Total Backlinks : {total}\n"
        f"✅ Good            : {good} ({pct_good}%)\n"
        f"⚠️ Manual Check    : {warn}\n"
        f"🚩 Flagged         : {flagged}\n"
    )

    if project_summary:
        manager_msg += f"\n<b>By Project:</b>\n"
        for proj, stats in sorted(project_summary.items()):
            manager_msg += (
                f"• <b>{proj}</b>: "
                f"✅{stats['good']} ⚠️{stats['warn']} 🚩{stats['flagged']}\n"
            )

    if flag_reasons:
        top_issues = sorted(flag_reasons.items(), key=lambda x: -x[1])[:6]
        manager_msg += f"\n<b>Top Issues:</b>\n"
        for reason, count in top_issues:
            manager_msg += f"• {count}x {reason}\n"

    if employee_data:
        manager_msg += f"\n<b>By Employee:</b>\n"
        for emp, stats in sorted(employee_data.items()):
            manager_msg += (
                f"• <b>{emp}</b>: "
                f"✅{stats['good']} ⚠️{stats['warn']} 🚩{stats['flagged']}\n"
            )

    manager_msg += f"\n⬇️ Download report from the platform"

    # ── Build team lead message (same as manager) ─────────────────────────
    teamlead_msg = manager_msg

    # ── Build per-employee messages ────────────────────────────────────────
    employee_msgs = {}
    for emp, stats in employee_data.items():
        emp_total   = stats["good"] + stats["flagged"] + stats["warn"]
        emp_pct     = round((stats["good"] / emp_total * 100)) if emp_total else 0
        msg = (
            f"👤 <b>Your Backlink Audit Report</b>\n"
            f"{'─' * 32}\n"
            f"Hi <b>{emp}</b>! Here's your audit summary:\n\n"
            f"📌 Total Submitted : {emp_total}\n"
            f"✅ Good            : {stats['good']} ({emp_pct}%)\n"
            f"⚠️ Manual Check    : {stats['warn']}\n"
            f"🚩 Needs Fixing    : {stats['flagged']}\n"
        )
        if stats["issues"]:
            msg += f"\n<b>🚩 Backlinks to Fix ({len(stats['issues'])}):</b>\n"
            for issue in stats["issues"][:8]:   # max 8 issues
                msg += (
                    f"\n• <b>{issue['activity']}</b>\n"
                    f"  🔗 {issue['url']}\n"
                    f"  🔑 KW: {issue['keyword']}\n"
                    f"  ❌ {issue['reason']}\n"
                )
        if stats["flagged"] == 0:
            msg += "\n🎉 Great work — all your backlinks passed!"
        msg += "\n\n📩 Contact your Team Lead if you have questions."
        employee_msgs[emp] = msg

    return {
        "manager_msg"  : manager_msg,
        "teamlead_msg" : teamlead_msg,
        "employee_msgs": employee_msgs,
    }

app = Flask(__name__)

# ── Optional password protection (set APP_PASSWORD on Render) ──────────────
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")


@app.before_request
def require_password():
    if not APP_PASSWORD or request.path == "/healthz":
        return None
    auth = request.authorization
    if auth and auth.password and hmac.compare_digest(auth.password, APP_PASSWORD):
        return None
    return Response("Password required", 401,
                    {"WWW-Authenticate": 'Basic realm="SEO Platform"'})


@app.route("/healthz")
def healthz():
    return "ok"


@app.route("/browser-check")
def browser_check():
    """Open this page to confirm headless Chrome can start on this server."""
    import subprocess, sys
    code = (
        "from playwright.sync_api import sync_playwright\n"
        "with sync_playwright() as p:\n"
        "    b = p.chromium.launch(headless=True, args=['--no-sandbox','--disable-dev-shm-usage'])\n"
        "    pg = b.new_page(); pg.set_content('<title>ok</title>'); print('CHROME OK —', b.version, '—', pg.title())\n"
        "    b.close()\n"
    )
    try:
        out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=90)
        text = (out.stdout + out.stderr).strip()
    except Exception as e:
        text = f"Check failed: {e}"
    ok = "CHROME OK" in text
    path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "(default)")
    body = (f"{'✅ Headless Chrome works' if ok else '❌ Headless Chrome did NOT start'}\n\n"
            f"PLAYWRIGHT_BROWSERS_PATH = {path}\n\n{text[-3000:]}")
    return Response(body, mimetype="text/plain; charset=utf-8")


# Allow ngrok to bypass the browser warning interstitial
@app.after_request
def add_headers(response):
    response.headers["ngrok-skip-browser-warning"] = "true"
    response.headers["Access-Control-Allow-Origin"] = "*"
    return response

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Store job status in memory
jobs = {}


# ── Serve the platform UI ──────────────────────────────────────────────────
@app.route("/")
def index():
    html_path = os.path.join(BASE_DIR, "seo_platform.html")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    from flask import Response
    return Response(content, mimetype="text/html")


# ── Detect SEO method from uploaded sheet ──────────────────────────────────
@app.route("/detect", methods=["POST"])
def detect():
    import traceback
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename.endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported"}), 400

    tmp_path = os.path.join(UPLOAD_FOLDER, f"_detect_{uuid.uuid4().hex[:6]}.xlsx")
    file.save(tmp_path)

    try:
        import pandas as pd
        xl         = pd.ExcelFile(tmp_path)
        all_sheets = xl.sheet_names
        if "Backlinks" in all_sheets:
            sheet_name = "Backlinks"
        elif "Worksheet" in all_sheets:
            sheet_name = "Worksheet"
        else:
            sheet_name = all_sheets[0]

        df = pd.read_excel(tmp_path, sheet_name=sheet_name, dtype=str).fillna("")

        total_rows = len(df)

        # Look for "Backlink Activity Type" column
        activity_col = None
        for col in df.columns:
            if "activity type" in col.lower() or "method" in col.lower():
                activity_col = col
                break

        # Which methods does our platform support?
        supported_methods = [
            "Blog Bookmarking",
            "Bookmarking",
            "Classifieds",
            "Web 2.0",
            "Article Submission",
            "Image Sharing",
            "Profile Creations",
            "Business Listing",
            "Press Release",
            "PDF Submission",
            "Guest Post",
            "Forums",
            "On-site Blog",
        ]

        methods = []   # list of dicts: {name, count, supported}
        primary_method   = "Blog Bookmarking"
        detected_any     = False

        if activity_col is not None:
            counts = df[activity_col].value_counts()
            for method_name, count in counts.items():
                cleaned = (method_name or "").strip()
                if not cleaned:
                    cleaned = "(blank)"
                methods.append({
                    "name"     : cleaned,
                    "count"    : int(count),
                    "supported": cleaned in supported_methods,
                })
            if methods:
                primary_method = methods[0]["name"]
                detected_any   = True

        if not methods:
            # Fallback if no activity column
            methods = [{
                "name"     : "Blog Bookmarking",
                "count"    : total_rows,
                "supported": True,
            }]

        is_mixed      = len(methods) > 1
        has_supported = any(m["supported"] for m in methods)

        # ── Extract project info from sheet ──────────────────────────────
        # Unique employees
        employees = []
        if "Employee Name" in df.columns:
            employees = [e for e in df["Employee Name"].dropna().unique().tolist()
                        if str(e).strip()]

        # Unique projects/websites
        projects = []
        if "Website" in df.columns:
            projects = [p for p in df["Website"].dropna().unique().tolist()
                       if str(p).strip()]

        # Project Manager (from column if exists, else blank)
        managers = []
        if "Project Manager" in df.columns:
            managers = [m for m in df["Project Manager"].dropna().unique().tolist()
                       if str(m).strip()]

        # Team Lead (from column if exists, else blank)
        team_leads = []
        if "Team Lead" in df.columns:
            team_leads = [t for t in df["Team Lead"].dropna().unique().tolist()
                         if str(t).strip()]

        # Project IDs
        project_ids = []
        if "Project ID" in df.columns:
            project_ids = [p for p in df["Project ID"].dropna().unique().tolist()
                          if str(p).strip()]

        # Re-save file with a stable filename so /audit can re-use it
        stable_id   = uuid.uuid4().hex[:8]
        stable_path = os.path.join(UPLOAD_FOLDER, f"detected_{stable_id}.xlsx")
        os.rename(tmp_path, stable_path)

        return jsonify({
            "methods"      : methods,
            "primary"      : primary_method,
            "detected"     : detected_any,
            "total_rows"   : total_rows,
            "is_mixed"     : is_mixed,
            "has_supported": has_supported,
            "supported_methods": supported_methods,
            "file_id"      : stable_id,
            "filename"     : file.filename,
            # Project info
            "employees"    : employees,
            "projects"     : projects,
            "managers"     : managers,
            "team_leads"   : team_leads,
            "project_ids"  : project_ids,
        })

    except Exception as e:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        error_detail = traceback.format_exc()
        print(f"DETECT ERROR: {error_detail}")
        return jsonify({"error": f"Could not read sheet: {str(e)}"}), 400


# ── Start audit job ────────────────────────────────────────────────────────
@app.route("/audit", methods=["POST"])
def audit():
    moz_key         = request.form.get("moz_key", "").strip()
    file_id         = request.form.get("file_id", "").strip()
    filter_methods  = request.form.get("filter_methods", "").strip()
    meta_project    = request.form.get("meta_project",    "").strip()
    meta_project_id = request.form.get("meta_project_id", "").strip()
    meta_manager    = request.form.get("meta_manager",    "").strip()
    meta_teamlead   = request.form.get("meta_teamlead",   "").strip()
    contacts_json   = request.form.get("contacts",        "[]")

    # Parse contacts
    import json as _json
    try:
        contacts = _json.loads(contacts_json)
    except Exception:
        contacts = []

    # Validate moz key format only if provided
    if moz_key and ":" not in moz_key:
        return jsonify({"error": "Invalid Moz key format. Use access_id:secret_key"}), 400

    # Two modes — file passed inline, OR file_id from a previous /detect call
    job_id      = str(uuid.uuid4())[:8]
    output_path = os.path.join(OUTPUT_FOLDER, f"{job_id}_AUDITED.xlsx")

    if file_id:
        input_path = os.path.join(UPLOAD_FOLDER, f"detected_{file_id}.xlsx")
        if not os.path.exists(input_path):
            return jsonify({"error": "Uploaded file expired — please re-upload"}), 400
    else:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
        file = request.files["file"]
        if not file.filename.endswith(".xlsx"):
            return jsonify({"error": "Only .xlsx files are supported"}), 400
        input_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_input.xlsx")
        file.save(input_path)

    jobs[job_id] = {
        "status"  : "running",
        "progress": 0,
        "total"   : 0,
        "good"    : 0,
        "flagged" : 0,
        "warn"    : 0,
        "logs"    : [],
        "output"  : output_path,
    }

    thread = threading.Thread(
        target=run_audit_job,
        args=(job_id, input_path, output_path, moz_key, filter_methods,
              meta_project, meta_project_id, meta_manager, meta_teamlead,
              contacts),
        daemon=True
    )
    thread.start()

    return jsonify({"job_id": job_id})


def run_audit_job(job_id, input_path, output_path, moz_key,
                  filter_methods="", meta_project="", meta_project_id="",
                  meta_manager="", meta_teamlead="", contacts=None):
    import subprocess, sys
    if contacts is None:
        contacts = []

    # Build dynamic recipient list from contacts
    # Merge with defaults — contacts from UI override defaults
    dynamic_ids = dict(DEFAULT_CHAT_IDS)  # start with defaults
    for c in contacts:
        tg = c.get("telegram", "").strip()
        if tg:
            role = c.get("role", "")
            dynamic_ids[tg] = role

    jobs[job_id]["logs"].append("🚀  Audit started...")
    send_telegram(
        f"🚀 <b>Audit Started</b>\n"
        f"🆔 Job: <code>{job_id}</code>\n"
        f"Will notify when complete.",
        dynamic_ids
    )

    cmd = [
        sys.executable, os.path.join(BASE_DIR, "backlink_auditor.py"),
        input_path,
        "--output", output_path,
    ]

    if moz_key:
        cmd += ["--moz-key", moz_key]
    if filter_methods and filter_methods.lower() != "all":
        cmd += ["--filter-methods", filter_methods]
    if meta_project:
        cmd += ["--meta-project", meta_project]
    if meta_project_id:
        cmd += ["--meta-project-id", meta_project_id]
    if meta_manager:
        cmd += ["--meta-manager", meta_manager]
    if meta_teamlead:
        cmd += ["--meta-teamlead", meta_teamlead]

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            cwd=BASE_DIR,
            env={**os.environ, "PYTHONUNBUFFERED": "1"},
        )

        for line in proc.stdout:
            line = line.rstrip()
            if not line:
                continue

            jobs[job_id]["logs"].append(line)

            # Parse progress from lines like [01/50]
            if "/" in line and "[" in line:
                try:
                    part = line.split("[")[1].split("]")[0]
                    cur, tot = part.split("/")
                    jobs[job_id]["progress"] = int(cur.strip())
                    jobs[job_id]["total"]    = int(tot.strip())
                except Exception:
                    pass

            # Parse result stats
            if "→ ✅" in line:
                jobs[job_id]["good"] += 1
            elif "→ 🚩" in line:
                jobs[job_id]["flagged"] += 1
            elif "→ ⚠️" in line:
                jobs[job_id]["warn"] += 1

        proc.wait()

        if proc.returncode == 0:
            jobs[job_id]["status"] = "done"
            jobs[job_id]["logs"].append("✅  Audit complete! Report is ready.")

            # ── Send detailed Telegram summaries ───────────────────────────
            try:
                filename = os.path.basename(jobs[job_id].get("output","audit"))
                summary  = build_telegram_summary(job_id, jobs, filename)

                # Send to manager
                mgr_ids = {c.get("telegram"): c.get("role","Manager")
                           for c in contacts if c.get("role") == "Manager" and c.get("telegram")}
                if not mgr_ids:
                    mgr_ids = {"702339188": "Nitish (Manager)"}
                send_telegram(summary["manager_msg"], mgr_ids)

                # Send to team lead
                tl_ids = {c.get("telegram"): c.get("role","Team Lead")
                          for c in contacts if c.get("role") == "Team Lead" and c.get("telegram")}
                if not tl_ids:
                    tl_ids = {"6595431574": "Tijeet (Team Lead)"}
                send_telegram(summary["teamlead_msg"], tl_ids)

                # Send per-employee messages
                emp_contacts = [c for c in contacts if c.get("role") == "Employee"]
                for emp, msg in summary["employee_msgs"].items():
                    sent = False
                    for ec in emp_contacts:
                        if ec.get("telegram"):
                            send_telegram(msg, {ec["telegram"]: emp})
                            sent = True
                            break
                    if not sent:
                        # Forward to team lead
                        send_telegram(
                            f"📩 <b>Report for {emp}</b> (not registered):\n\n{msg}",
                            tl_ids
                        )

                # Send emails if provided
                emails = [(c.get("email","").strip(), c.get("role","")) 
                          for c in contacts if c.get("email","").strip()]
                if emails:
                    jobs[job_id]["logs"].append(
                        f"📧 Sending emails to: {', '.join(e for e, _ in emails)}"
                    )
                    for email_addr, role in emails:
                        if role == "Manager":
                            email_body = summary["manager_msg"]
                        elif role == "Team Lead":
                            email_body = summary["teamlead_msg"]
                        else:
                            # Find matching employee msg
                            email_body = next(
                                (msg for emp, msg in summary["employee_msgs"].items()),
                                summary["teamlead_msg"]
                            )
                        # Strip HTML tags for email subject
                        subject = f"Backlink Audit Report — Job {job_id}"
                        # Clean HTML tags from body
                        import re as _re
                        clean_body = _re.sub(r'<[^>]+>', '', email_body)
                        send_email(email_addr, subject, clean_body)

            except Exception as e:
                print(f"Telegram notify error: {e}")
        else:
            jobs[job_id]["status"] = "error"
            jobs[job_id]["logs"].append("❌  Audit failed. Check file format.")
            send_telegram(
                f"❌ <b>Audit Failed</b>\n"
                f"🆔 Job: <code>{job_id}</code>\n"
                f"Please check the file format and try again.",
                DEFAULT_CHAT_IDS
            )

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["logs"].append(f"❌  Server error: {str(e)}")


# ── Poll job status ────────────────────────────────────────────────────────
@app.route("/status/<job_id>")
def status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    return jsonify({
        "status"  : job["status"],
        "progress": job["progress"],
        "total"   : job["total"],
        "good"    : job["good"],
        "flagged" : job["flagged"],
        "warn"    : job["warn"],
        "logs"    : job["logs"][-30:],
    })


# ── Download result ────────────────────────────────────────────────────────
@app.route("/download/<job_id>")
def download(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "done":
        return jsonify({"error": "Audit not finished yet"}), 400
    if not os.path.exists(job["output"]):
        return jsonify({"error": "Output file not found"}), 404
    return send_file(
        job["output"],
        as_attachment=True,
        download_name="Backlinks_AUDITED.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ── Run ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import socket
    # PORT env var wins (Render sets it); else try 8080, then 5000
    port = int(os.environ.get("PORT", 0)) or 8080
    if "PORT" not in os.environ:
        s = socket.socket()
        try:
            s.bind(("", port))
            s.close()
        except OSError:
            port = 5000

    print("\n" + "═" * 52)
    print("   Backlink Auditor v3 — Web UI Server")
    print("═" * 52)
    print(f"   Open in browser:  http://localhost:{port}")
    print(f"   For ngrok run:     ngrok http {port}")
    print("   Press Ctrl+C to stop")
    print("═" * 52 + "\n")
    app.run(debug=False, host="0.0.0.0", port=port)

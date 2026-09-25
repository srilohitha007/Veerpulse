"""Authenticated route audit using Flask test client sessions. No DB writes except create_app create_all no-op."""
from __future__ import annotations

import traceback

from sqlalchemy import text
from app import create_app, db, Personnel, WelfareOfficer

app = create_app()
client = app.test_client()


def get_ids():
    with app.app_context():
        personnel = [p.personnel_id for p in Personnel.query.order_by(Personnel.personnel_id).all()]
        officers = [o.officer_id for o in WelfareOfficer.query.order_by(WelfareOfficer.officer_id).all()]
        print("personnel_ids", personnel)
        print("officer_ids", officers)
        return personnel, officers


def as_role(role, **session_vals):
    with client.session_transaction() as sess:
        sess.clear()
        sess["role"] = role
        sess.update(session_vals)


def hit(method, path):
    try:
        resp = client.open(path, method=method, follow_redirects=False)
        loc = resp.headers.get("Location", "")
        body = resp.get_data(as_text=True)
        err = ""
        if resp.status_code >= 400:
            err = body[:180].replace("\n", " ")
        elif "Traceback" in body or "Internal Server Error" in body:
            err = "traceback in body"
        print(f"{resp.status_code} {method} {path} loc={loc} {err}")
        return resp.status_code
    except Exception as exc:
        print(f"EXC {method} {path} {exc}")
        traceback.print_exc()
        return 0


def main():
    personnel, officers = get_ids()
    pid = personnel[0] if personnel else "MISSING"
    oid = officers[0] if officers else "MISSING"

    personnel_gets = [
        "/personnel-dashboard",
        "/personnel-profile",
        "/personnel-duty-leave",
        "/personnel-trends",
        "/personnel-analysis",
        "/personnel-wellness-checkin",
        "/personnel-wellness-scan",
        "/personnel-support",
        "/personnel-notifications",
        "/personnel-privacy",
        "/api/personnel/analysis",
        "/api/personnel/current-assessment",
        "/api/personnel/operational-summary",
        "/api/personnel/personal-baseline",
        "/api/personnel/leave-requests",
        "/api/personnel/leave-metrics",
        "/api/personnel/wellness-checkins",
    ]
    officer_gets = [
        "/welfare-officer-dashboard",
        "/welfare-officer-personnel-overview",
        "/welfare-officer-leave-management",
        "/welfare-officer-alerts",
        "/welfare-officer-support-requests",
        "/welfare-officer-case-review",
        f"/welfare-officer-case-review/{pid}",
        "/welfare-officer-case-review/DOES-NOT-EXIST",
        "/welfare-officer-priority-review",
        "/welfare-officer-operational-review",
        f"/welfare-officer-operational-review?personnel_id={pid}",
        "/welfare-officer-intervention-recommendations",
        f"/welfare-officer-intervention-recommendations?personnel_id={pid}",
        "/welfare-officer-unit-insights",
        "/welfare-officer-hrms-integration",
        "/welfare-officer-duty-upload",
        "/welfare-officer-reports",
        f"/welfare-officer-reports?personnel_id={pid}",
        "/welfare-officer-notifications",
        "/welfare-officer-privacy-audit",
        "/welfare-officer-settings",
        f"/welfare-officer-analysis/{pid}",
        "/welfare-officer-analysis/DOES-NOT-EXIST",
        "/api/welfare-officer/personnel",
        f"/api/welfare-officer/personnel/{pid}",
        f"/api/welfare-officer/personnel/{pid}/analysis",
        f"/api/welfare-officer/personnel/{pid}/leave-metrics",
        "/api/welfare-officer/leave-requests",
        "/api/welfare-officer/unit-insights",
        "/api/welfare-officer/audit-events",
        f"/api/welfare-officer/recovery-window-analysis?personnel_id={pid}",
        f"/api/welfare-officer/welfare-actions?personnel_id={pid}",
        "/api/welfare-officer/hrms/template",
        "/api/welfare-officer/personnel/bulk-template",
    ]

    print("--- personnel session ---")
    as_role("personnel", personnel_id=pid)
    for path in personnel_gets:
        hit("GET", path)

    print("--- personnel IDOR attempt ---")
    other = next((p for p in personnel if p != pid), None)
    if other:
        hit("GET", f"/api/welfare-officer/personnel/{other}")
        # personnel APIs should ignore query personnel_id
        hit("GET", f"/api/personnel/analysis?personnel_id={other}")
        hit("GET", f"/api/personnel/current-assessment?personnel_id={other}")

    print("--- officer session ---")
    as_role("welfare_officer", officer_id=oid)
    for path in officer_gets:
        hit("GET", path)

    print("--- officer hitting personnel page ---")
    hit("GET", "/personnel-dashboard")
    hit("POST", "/save-wellness-scan")

    print("--- personnel hitting officer page ---")
    as_role("personnel", personnel_id=pid)
    hit("GET", "/welfare-officer-dashboard")
    hit("GET", f"/welfare-officer-case-review/{pid}")
    hit("GET", "/api/welfare-officer/personnel")


if __name__ == "__main__":
    main()

"""Temporary read-only runtime audit. Do not commit secrets."""
from __future__ import annotations

import os
import traceback

from dotenv import load_dotenv
from sqlalchemy import create_engine, inspect, text

load_dotenv()


def main() -> None:
    engine = create_engine(os.getenv("DATABASE_URL"))
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
        print("postgres_connect OK")

    insp = inspect(engine)
    tables = sorted(insp.get_table_names())
    print("table_count", len(tables))
    print("tables", ",".join(tables))

    with engine.connect() as conn:
        for table in tables:
            cols = [c["name"] for c in insp.get_columns(table)]
            n = conn.execute(text(f'SELECT COUNT(*) FROM "{table}"')).scalar()
            print(f"table {table} cols={len(cols)} rows={n}")

    from app import create_app

    app = create_app()
    print("create_app OK")
    print("route_count", len(list(app.url_map.iter_rules())))

    client = app.test_client()
    public_paths = [
        "/",
        "/personnel-login",
        "/welfare-officer-login",
        "/personnel-dashboard",
        "/welfare-officer-dashboard",
        "/logout",
        "/save-wellness-scan",
        "/api/personnel/analysis",
        "/welfare-officer-case-review",
        "/welfare-officer-case-review/NOBODY",
        "/static/css/style.css",
    ]
    for path in public_paths:
        try:
            method = "POST" if path == "/save-wellness-scan" else "GET"
            resp = client.open(path, method=method)
            print(f"anon {method} {path} -> {resp.status_code}")
        except Exception:
            print(f"anon FAIL {path}")
            traceback.print_exc()

    # Template compile check
    with app.app_context():
        from flask import render_template
        templates = [
            "home.html",
            "personnel_login.html",
            "welfare_officer_login.html",
        ]
        for name in templates:
            render_template(name)
            print("render_ok", name)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise

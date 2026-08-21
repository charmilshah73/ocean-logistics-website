"""Production server for IIS reverse-proxy (Waitress).

Usage (IIS / local production):
  python serve_production.py

Optional env:
  OCEAN_HOST   default 127.0.0.1 (IIS proxies here; use 0.0.0.0 only if exposing Python directly)
  OCEAN_PORT   default 8050
"""
from __future__ import annotations

import os

from app.server import app, bootstrap_data


def main() -> None:
    bootstrap_data()
    from waitress import serve

    host = os.environ.get("OCEAN_HOST", "127.0.0.1")
    port = int(os.environ.get("OCEAN_PORT", "8050"))
    print(f"Ocean Logistics (Waitress) on http://{host}:{port}")
    print(f"Dashboard: http://127.0.0.1:{port}/")
    print(f"Admin:     http://127.0.0.1:{port}/admin")
    serve(
        app,
        host=host,
        port=port,
        threads=8,
        channel_timeout=300,
        connection_limit=200,
        max_request_body_size=100 * 1024 * 1024,
    )


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
import argparse
import json
import os
from pathlib import Path

import requests

def fetch_system(host: str, username: str, password: str, verify_tls: bool = True):
    url = host.rstrip("/") + "/redfish/v1/Systems/System.Embedded.1"
    response = requests.get(url, auth=(username, password), verify=verify_tls, timeout=20)
    response.raise_for_status()
    return response.json()

def main():
    parser = argparse.ArgumentParser(description="Fetch a Redfish system JSON payload.")
    parser.add_argument("--host", default=os.getenv("REDFISH_HOST"))
    parser.add_argument("--username", default=os.getenv("REDFISH_USERNAME"))
    parser.add_argument("--password", default=os.getenv("REDFISH_PASSWORD"))
    parser.add_argument("--out", type=Path, default=Path("reports/redfish-system.json"))
    parser.add_argument("--insecure", action="store_true", help="Disable TLS certificate validation.")
    args = parser.parse_args()

    if not args.host or not args.username or not args.password:
        raise SystemExit("Missing host/username/password. Use flags or environment variables.")

    args.out.parent.mkdir(parents=True, exist_ok=True)
    data = fetch_system(args.host, args.username, args.password, verify_tls=not args.insecure)
    args.out.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Wrote {args.out}")

if __name__ == "__main__":
    main()

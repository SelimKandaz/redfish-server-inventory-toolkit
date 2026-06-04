#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

def main():
    parser = argparse.ArgumentParser(description="Generate a server inventory report from fake Redfish JSON.")
    parser.add_argument("json_file", type=Path)
    parser.add_argument("--out", type=Path, default=Path("reports"))
    args = parser.parse_args()

    args.out.mkdir(parents=True, exist_ok=True)
    data = json.loads(args.json_file.read_text(encoding="utf-8"))

    lines = [
        "# Server Inventory Report",
        "",
        f"Model: {data.get('Model', 'Unknown')}",
        f"Serial: {data.get('SerialNumber', 'Unknown')}",
        "",
        "## CPU",
        f"- Count: {data.get('ProcessorSummary', {}).get('Count', 'Unknown')}",
        f"- Model: {data.get('ProcessorSummary', {}).get('Model', 'Unknown')}",
        "",
        "## Memory",
        f"- Total GiB: {data.get('MemorySummary', {}).get('TotalSystemMemoryGiB', 'Unknown')}",
        "",
        "## Storage",
        f"- Controller: {data.get('StorageSummary', {}).get('Controller', 'Unknown')}",
        f"- Drive Count: {data.get('StorageSummary', {}).get('DriveCount', 'Unknown')}",
        "",
        "## Network",
        f"- Adapters: {data.get('NetworkSummary', {}).get('AdapterCount', 'Unknown')}",
        f"- Ports: {data.get('NetworkSummary', {}).get('PortCount', 'Unknown')}",
    ]
    out = args.out / "server-inventory-report.md"
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Generated {out}")

if __name__ == "__main__":
    main()

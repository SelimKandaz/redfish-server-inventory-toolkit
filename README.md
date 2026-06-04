# Redfish Server Inventory Toolkit

A sanitized public toolkit for collecting server inventory through Redfish-style APIs and generating structured hardware reports.

This repository uses fake sample responses by default. It does not include real server IPs, credentials, service tags or customer inventory data.

## Features

- Parse fake Redfish inventory responses
- Normalize server model, serial, CPU, memory and storage data
- Generate Markdown inventory summaries
- Include an optional Redfish fetch script using environment variables
- Provide a safe public example of server inventory automation

## Quick start with sample data

```bash
python src/redfish_inventory_demo.py examples/redfish-system-demo.json --out reports
```

## Optional live Redfish fetch

Use only on systems you own or are authorized to access.

```bash
pip install -r requirements.txt
python src/redfish_fetch_system.py --host https://192.0.2.10 --username demo_user --password change_me --out reports/redfish-system.json --insecure
```

Do not commit real credentials or real inventory output.

# Redfish Server Inventory Toolkit

A sanitized public toolkit concept for collecting server inventory through Redfish-style APIs and generating structured hardware reports.

This repository uses fake sample responses only. It does not include real server IPs, credentials, service tags or customer inventory data.

## Features

- Parse fake Redfish inventory responses
- Normalize server model, serial, CPU, memory and storage data
- Generate Markdown inventory summaries
- Provide a safe public example of server inventory automation

## Quick start

```bash
python src/redfish_inventory_demo.py examples/redfish-system-demo.json --out reports
```

## Technology focus

- Python
- Redfish API concepts
- Server hardware inventory
- JSON parsing
- Markdown report generation
- Infrastructure operations automation

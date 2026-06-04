# Redfish Server Inventory Toolkit

An open-source Redfish inventory toolkit for collecting server hardware data and generating structured reports.

This repository now includes:

1. A simple Redfish JSON report demo.
2. An optional live Redfish fetch script.
3. A sanitized desktop configurator application based on the original working tool.

## Simple sample-data report

```bash
python src/redfish_inventory_demo.py examples/redfish-system-demo.json --out reports
```

## Optional live Redfish fetch

Use only on systems you own or are authorized to access.

```bash
pip install -r requirements.txt
python src/redfish_fetch_system.py --host https://192.0.2.10 --username demo_user --password change_me --out reports/redfish-system.json --insecure
```

## Desktop configurator app

```bash
pip install -r requirements.txt
python run_desktop.py
```

Expected local files for the desktop app:

- `task.csv`, same structure as `examples/task.example.csv`
- a fillable PDF template if you want PDF output

Do not commit your real `task.csv`, credentials or generated inventory reports.

## License

MIT

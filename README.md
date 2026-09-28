# Redfish Server Inventory Toolkit

Collects server hardware data over Redfish and turns it into structured reports.

## Report from a Redfish JSON file

```bash
python src/redfish_inventory_demo.py examples/redfish-system-demo.json --out reports
```

## Fetch live from a server

```bash
pip install -r requirements.txt
python src/redfish_fetch_system.py --host https://192.0.2.10 --username demo_user --password change_me --out reports/redfish-system.json --insecure
```

Only use it on systems you are authorized to access.

## Desktop configurator

```bash
python run_desktop.py
```

Needs a `task.csv` (see `examples/task.example.csv`) and, for PDF output, a fillable PDF template.

## License

MIT

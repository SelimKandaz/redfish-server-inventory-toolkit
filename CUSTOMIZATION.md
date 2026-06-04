# Customization Guide

This project can be adapted for server inventory and Redfish-based reporting workflows.

## Common customization points

- Redfish endpoint path
- Authentication method
- TLS verification policy
- Vendor-specific JSON fields
- Report template
- CSV task file format
- Desktop workflow labels

## Start here

- `src/redfish_inventory_demo.py` for sample JSON reporting
- `src/redfish_fetch_system.py` for live Redfish fetches
- `run_desktop.py` for the desktop configurator app
- `examples/task.example.csv` for task CSV structure

## Keep local

Keep real IP lists, credentials and generated reports outside the repository.

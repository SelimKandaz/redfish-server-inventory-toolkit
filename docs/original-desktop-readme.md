# Server Configurator Desktop

A PySide6 desktop front end for the old unified Redfish CLI collector.

## What it does
- Opens as a normal Windows app
- Loads `task.csv` on startup
- Lets you add, edit, delete, enable, disable, and save targets
- Runs Redfish collection in a background worker so the UI stays responsive
- Reuses the same PDF and Excel generation flow from the old script
- Keeps `task.csv` in the same `IP,User,Password` format

## Expected files in the same folder as `run_desktop.py`
- `task.csv`
- `Server Configurator_fillable.pdf`

## Install
```bash
pip install -r requirements.txt
```

## Run
```bash
python run_desktop.py
```

## Notes
- This is the first desktop version. It is structured so we can keep improving it.
- The app does not change the CSV schema. It still saves three columns only: `IP,User,Password`.
- Row enable/disable state is a UI-only state for now.

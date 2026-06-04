# -*- coding: utf-8 -*-
"""
Unified Server Configurator (CLI)
- Dell iDRAC / HPE iLO / Supermicro BMC (Redfish)
- Fills "Server Configurator_fillable.pdf"
- Builds Excel with serial lists
- Groups multiple servers: if configs match -> only Quantity increases;
  if not -> writes per-IP differences into EXTRAS.

Requirements (pip):
  pip install requests openpyxl PyPDF2

Files expected in same folder:
  - task.csv                 (columns: IP,User,Password)
  - Server Configurator_fillable.pdf
  - This script

Output structure:
  Configurator_PDFs_<Vendor>/
      <RUN or SO# folder>/
          Configurator_<IP>.pdf
          Configurator_GROUP.pdf
          <SO#...>_serials.xlsx
"""

import os
import csv
from collections import Counter
from datetime import datetime

import requests
import urllib3
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject
from openpyxl import Workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PDF = os.path.join(BASE_DIR, "Server Configurator_fillable.pdf")

DEFAULT_OUTPUT_DIRS = {
    "dell": os.path.join(BASE_DIR, "Configurator_PDFs_Dell"),
    "hpe": os.path.join(BASE_DIR, "Configurator_PDFs"),
    "supermicro": os.path.join(BASE_DIR, "Configurator_PDFs_Supermicro"),
}
for d in DEFAULT_OUTPUT_DIRS.values():
    os.makedirs(d, exist_ok=True)

# --------------------------------------------------------------------
# PDF field map – same as working Dell script
# --------------------------------------------------------------------
FIELD_MAP = {
    "CHASSIS_PLATFORM": "Text4",
    "QUANTITY": "Text5",
    "MOTHERBOARD": "Text6",
    "SYSTEM_SERIALS": "Text8",
    "CPU": "Text9",
    "MEMORY": "Text10",
    "DRIVES": "Text12",
    "RAID_HW": "Text16",
    "NETWORK_HW": "Text17",
    "PSU": "Text19",
    "EXTRAS": "Text20",
    "BIOS": "Text21",
    "BMC_ILO": "Text22",
}

# --------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------
def rf_get(ip, path, auth, timeout=15):
    """Simple Redfish GET with basic error handling."""
    if path.startswith("http"):
        url = path
    else:
        url = f"https://{ip}{path}"
    try:
        r = requests.get(url, auth=auth, verify=False, timeout=timeout)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[{ip}] GET {path} failed: {e}")
        return {}

def rf_first_member_path(collection):
    if not collection:
        return None
    members = collection.get("Members") or []
    if not members:
        return None
    return members[0].get("@odata.id")

# --------------------------------------------------------------------
# Path discovery
# --------------------------------------------------------------------
def discover_paths(ip, auth, vendor_key):
    # Systems
    sys_coll = rf_get(ip, "/redfish/v1/Systems", auth)
    sys_path = rf_first_member_path(sys_coll)
    if not sys_path:
        if vendor_key == "dell":
            sys_path = "/redfish/v1/Systems/System.Embedded.1"
        else:
            sys_path = "/redfish/v1/Systems/1"

    # Chassis
    ch_coll = rf_get(ip, "/redfish/v1/Chassis", auth)
    ch_path = rf_first_member_path(ch_coll)
    if not ch_path:
        if vendor_key == "dell":
            ch_path = "/redfish/v1/Chassis/System.Embedded.1"
        else:
            ch_path = "/redfish/v1/Chassis/1"

    # Managers
    mgr_coll = rf_get(ip, "/redfish/v1/Managers", auth)
    mgr_path = rf_first_member_path(mgr_coll) or "/redfish/v1/Managers/1"

    return sys_path, ch_path, mgr_path

# --------------------------------------------------------------------
# Collectors
# --------------------------------------------------------------------
def get_system(ip, auth, sys_path):
    return rf_get(ip, sys_path, auth)

def get_chassis(ip, auth, ch_path):
    return rf_get(ip, ch_path, auth)

def get_bios_version(system):
    for key in ("BiosVersion", "BIOSVersion", "Bios"):
        v = system.get(key)
        if isinstance(v, str) and v:
            return v
        if isinstance(v, dict):
            for k2 in ("Current", "Version", "VersionString"):
                if isinstance(v.get(k2), str) and v.get(k2):
                    return v.get(k2)
    oem = system.get("Oem") or {}
    for v in oem.values():
        if isinstance(v, dict):
            for k2, v2 in v.items():
                if isinstance(v2, str) and "bios" in k2.lower():
                    return v2
    return ""

def get_bmc_version(ip, auth, mgr_path):
    mgr = rf_get(ip, mgr_path, auth)
    if not mgr:
        return ""
    for key in ("FirmwareVersion", "ManagerFirmwareVersion", "Version"):
        if isinstance(mgr.get(key), str) and mgr.get(key):
            return mgr.get(key)
    oem = mgr.get("Oem") or {}
    for v in oem.values():
        if isinstance(v, dict):
            for k2, v2 in v.items():
                if isinstance(v2, str) and ("fw" in k2.lower() or "firmware" in k2.lower()):
                    return v2
    return ""

def get_processors(ip, auth, sys_path):
    out = []
    coll = rf_get(ip, f"{sys_path}/Processors", auth)
    for m in (coll.get("Members") or []):
        p = rf_get(ip, m.get("@odata.id"), auth)
        if p:
            out.append(p)
    return out

def get_dimms(ip, auth, sys_path):
    out = []
    coll = rf_get(ip, f"{sys_path}/Memory", auth)
    for m in (coll.get("Members") or []):
        d = rf_get(ip, m.get("@odata.id"), auth)
        if d:
            out.append(d)
    return out

def get_drives(ip, auth, sys_path, ch_path):
    """
    Collect drive-like devices (HDD, SSD, NVMe, including many M.2 devices).

    Strategy:
      - Use System -> Storage -> Drives (standard Redfish)
      - Also look at Chassis-level /Drives if present
      - On some Supermicro BMCs, NVMe M.2 devices only appear under
        Chassis -> PCIeDevices, so we treat storage-like PCIe devices
        as drives as well.
    """
    out = []

    def add_drive_like(d):
        if not d:
            return
        sn = d.get("SerialNumber")

        # Try common Redfish Identifier patterns for serial
        if not sn:
            for ident in d.get("Identifiers", []) or []:
                if isinstance(ident, dict):
                    fmt = (ident.get("DurableNameFormat") or "").lower()
                    if "serial" in fmt or "sn" in fmt:
                        sn = ident.get("DurableName")
                        if sn:
                            break

        # Try OEM nested fields for any key that looks like a serial
        if not sn:
            oem = d.get("Oem") or {}
            for v in oem.values():
                if isinstance(v, dict):
                    for k2, v2 in v.items():
                        if isinstance(v2, str) and "serial" in k2.lower():
                            sn = v2
                            break
                if sn:
                    break

        d = dict(d)
        d["_BetterSerial"] = sn
        out.append(d)

    # 1) System-level Storage -> Drives (most vendors)
    store_coll = rf_get(ip, f"{sys_path}/Storage", auth)
    for m in (store_coll.get("Members") or []):
        s = rf_get(ip, m.get("@odata.id"), auth)
        for dref in (s.get("Drives") or []):
            d = rf_get(ip, dref.get("@odata.id"), auth)
            add_drive_like(d)

    # If no chassis info, return what we have
    if not ch_path:
        return out

    # 2) Chassis-level Drives collections (used by some Supermicro BMCs)
    ch_obj = rf_get(ip, ch_path, auth)
    for dref in (ch_obj.get("Drives") or []):
        d = rf_get(ip, dref.get("@odata.id"), auth)
        add_drive_like(d)

    ch_drives_coll = rf_get(ip, f"{ch_path}/Drives", auth)
    for m in (ch_drives_coll.get("Members") or []):
        d = rf_get(ip, m.get("@odata.id"), auth)
        add_drive_like(d)

    # 3) Storage-like PCIe devices (M.2 NVMe on some platforms)
    pcie_coll = rf_get(ip, f"{ch_path}/PCIeDevices", auth)
    for m in (pcie_coll.get("Members") or []):
        pcie = rf_get(ip, m.get("@odata.id"), auth)
        if not pcie:
            continue

        dev_type = (pcie.get("DeviceType") or "").lower()
        model = (pcie.get("Model") or pcie.get("Name") or "").lower()

        # Filter to likely storage devices
        if not any(x in dev_type for x in ["nvme", "ssd", "storage"]):
            if "nvme" not in model and "ssd" not in model:
                continue

        d = {
            "Model": pcie.get("Model") or pcie.get("Name"),
            "PartNumber": pcie.get("PartNumber"),
            "SerialNumber": pcie.get("SerialNumber"),
            "CapacityBytes": pcie.get("CapacityBytes"),
            "Oem": pcie.get("Oem"),
            "Identifiers": pcie.get("Identifiers"),
        }
        add_drive_like(d)

    return out

def get_nic_ports(ip, auth, sys_path, mgr_path):
    """EthernetInterfaces – used only for PORT info and MAC addresses."""
    out = []
    coll = rf_get(ip, f"{sys_path}/EthernetInterfaces", auth)
    for m in (coll.get("Members") or []):
        n = rf_get(ip, m.get("@odata.id"), auth)
        if n:
            out.append(n)
    if out:
        return out
    coll2 = rf_get(ip, f"{mgr_path}/EthernetInterfaces", auth)
    for m in (coll2.get("Members") or []):
        n = rf_get(ip, m.get("@odata.id"), auth)
        if n:
            out.append(n)
    return out

def get_nic_adapters(ip, auth, sys_path, ch_path):
    """
    Try to read FRU-level NetworkAdapters (for NIC card serial numbers).

    Many Supermicro BMCs expose:
        /redfish/v1/Chassis/1/NetworkAdapters
    and sometimes also under Systems.
    """
    out = []

    # 1) Chassis-level NetworkAdapters
    coll = rf_get(ip, f"{ch_path}/NetworkAdapters", auth)
    for m in (coll.get("Members") or []):
        a = rf_get(ip, m.get("@odata.id"), auth)
        if a:
            out.append(a)

    # 2) System-level NetworkAdapters (fallback)
    coll2 = rf_get(ip, f"{sys_path}/NetworkAdapters", auth)
    for m in (coll2.get("Members") or []):
        a = rf_get(ip, m.get("@odata.id"), auth)
        if a:
            out.append(a)

    return out

def get_psus(ip, auth, ch_path):
    out = []
    power = rf_get(ip, f"{ch_path}/Power", auth)
    for p in (power.get("PowerSupplies") or []):
        if not isinstance(p, dict):
            continue
        sn = p.get("SerialNumber")
        if not sn:
            oem = p.get("Oem") or {}
            for v in oem.values():
                if isinstance(v, dict):
                    for k2, v2 in v.items():
                        if isinstance(v2, str) and "serial" in k2.lower():
                            sn = v2
                            break
                if sn:
                    break
        out.append({
            "SerialNumber": sn,
            "Model": p.get("Model"),
            "CapacityWatts": p.get("CapacityWatts") or p.get("PowerCapacityWatts"),
            "LineInputVoltageType": p.get("LineInputVoltageType"),
        })
    return out

# --------------------------------------------------------------------
# Summaries for PDF
# --------------------------------------------------------------------
def summarize_cpu(procs):
    if not procs:
        return ""
    # Group by (Model, PartNumber) so CPU PN also appears in PDF text
    groups = Counter()
    for p in procs:
        model = p.get("Model") or ""
        pn = p.get("PartNumber") or ""
        key = (model, pn)
        groups[key] += 1
    parts = []
    for (model, pn), count in groups.items():
        if not model and not pn:
            continue
        base = ""
        if count > 1:
            base += f"{count}x "
        if model:
            base += model
        if pn:
            if model:
                base += f" ({pn})"
            else:
                base += pn
        parts.append(base)
    return "; ".join(parts)

def summarize_memory(dimms):
    if not dimms:
        return ""
    total_mib = 0
    by_pn = Counter()
    sample_size = {}
    for d in dimms:
        cap = d.get("CapacityMiB") or 0
        total_mib += cap
        pn = d.get("PartNumber") or d.get("Manufacturer") or "Unknown"
        by_pn[pn] += 1
        if pn not in sample_size and d.get("CapacityMiB"):
            sample_size[pn] = int(d.get("CapacityMiB")) // 1024
    total_gb = round(total_mib / 1024.0, 1)
    parts = [f"Total: {total_gb}GB"]
    for pn, qty in by_pn.items():
        sz = sample_size.get(pn)
        if sz:
            parts.append(f"{qty}x {sz}GB ({pn})")
        else:
            parts.append(f"{qty}x ({pn})")
    return "; ".join(parts)

def summarize_psus(psus):
    if not psus:
        return ""
    groups = Counter()
    for p in psus:
        watts = p.get("CapacityWatts")
        volt = p.get("LineInputVoltageType")
        model = p.get("Model")
        key = (watts, volt, model)
        groups[key] += 1
    parts = []
    for (watts, volt, model), count in groups.items():
        txt = ""
        if count and count > 1:
            txt += f"{count}x "
        if watts:
            txt += f"{watts}W"
        if volt:
            txt += f" {volt}"
        if model:
            if txt:
                txt += f" ({model})"
            else:
                txt = model
        parts.append(txt.strip())
    return "; ".join(parts)

def summarize_drives(drives):
    if not drives:
        return ""
    groups = Counter()
    for d in drives:
        model = d.get("Model") or "Drive"
        pn = d.get("PartNumber") or ""
        size = d.get("CapacityBytes")
        sizegb = None
        try:
            if isinstance(size, int) and size > 0:
                sizegb = int(round(size / (1024 ** 3)))
        except Exception:
            sizegb = None
        key = (pn, model, sizegb)
        groups[key] += 1
    parts = []
    for (pn, model, sizegb), qty in groups.items():
        label_model = model or "Drive"
        if sizegb:
            base = f"{qty}x {sizegb}GB {label_model}"
        else:
            base = f"{qty}x {label_model}"
        if pn:
            base += f" ({pn})"
        parts.append(base)
    return "; ".join(parts)

def summarize_network(nic_adapters):
    """Summarize NIC cards for PDF using FRU-level adapters.

    Prefer Model + PartNumber for each physical NIC card.
    For Supermicro BMC, PartNumber may live under OEM fields, so we check there.
    """
    if not nic_adapters:
        return ""
    groups = Counter()
    for a in nic_adapters:
        model = a.get("Model") or a.get("Name") or "NIC"

        # Try clean part number
        pn = a.get("PartNumber") or ""
        if not pn:
            oem = a.get("Oem") or {}
            for v in oem.values():
                if isinstance(v, dict):
                    for k2, v2 in v.items():
                        if isinstance(v2, str) and ("part" in k2.lower() or "spn" in k2.lower()):
                            pn = v2
                            break
                if pn:
                    break

        key = (model, pn)
        groups[key] += 1
    parts = []
    for (model, pn), qty in groups.items():
        base = ""
        if qty > 1:
            base += f"{qty}x "
        base += model
        if pn:
            base += f" ({pn})"
        parts.append(base)
    return "; ".join(parts)

def write_pdf_from_values(ip, values, output_pdf):
    """
    values: dict with logical keys from FIELD_MAP (CHASSIS_PLATFORM, CPU, ...)
    """
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError(f"Template PDF bulunamadı: {TEMPLATE_PDF}")

    fields = {
        FIELD_MAP["CHASSIS_PLATFORM"]: values.get("CHASSIS_PLATFORM", ""),
        FIELD_MAP["QUANTITY"]: values.get("QUANTITY", ""),
        FIELD_MAP["MOTHERBOARD"]: values.get("MOTHERBOARD", ""),
        FIELD_MAP["SYSTEM_SERIALS"]: values.get("SYSTEM_SERIALS", ""),
        FIELD_MAP["CPU"]: values.get("CPU", ""),
        FIELD_MAP["MEMORY"]: values.get("MEMORY", ""),
        FIELD_MAP["DRIVES"]: values.get("DRIVES", ""),
        FIELD_MAP["RAID_HW"]: values.get("RAID_HW", ""),
        FIELD_MAP["NETWORK_HW"]: values.get("NETWORK_HW", ""),
        FIELD_MAP["PSU"]: values.get("PSU", ""),
        FIELD_MAP["EXTRAS"]: values.get("EXTRAS", ""),
        FIELD_MAP["BIOS"]: values.get("BIOS", ""),
        FIELD_MAP["BMC_ILO"]: values.get("BMC_ILO", ""),
    }

    reader = PdfReader(TEMPLATE_PDF)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)
    # Copy AcroForm and set NeedAppearances so fields show in viewers
    try:
        if '/AcroForm' in reader.trailer['/Root']:
            writer._root_object.update({NameObject('/AcroForm'): reader.trailer['/Root']['/AcroForm']})
            writer._root_object['/AcroForm'].update({NameObject('/NeedAppearances'): BooleanObject(True)})
    except Exception:
        pass

    for page in writer.pages:
        writer.update_page_form_field_values(page, fields)

    if reader.metadata:
        try:
            writer.add_metadata(reader.metadata)
        except Exception:
            pass

    with open(output_pdf, "wb") as f:
        writer.write(f)

# --------------------------------------------------------------------
# Excel serials
# --------------------------------------------------------------------
def build_serial_columns(servers):
    headers = [
        "Chassis SN:",
        "MB SN:",
        "Blade SN:",
        "CPU SN:",
        "Memory SN:",
        "NIC card SN:",
        "NIC MAC:",
        "Storage SN:",
        "PSU SN:",
        "Server SN:",
    ]
    cols = {h: [] for h in headers}

    for s in servers:
        vendor = (s.get("vendor") or "").lower()
        system = s["system"] or {}
        chassis = s["chassis"] or {}
        procs = s["procs"] or []
        dimms = s["dimms"] or []
        drives = s["drives"] or []
        nic_ports = s["nic_ports"] or []
        nic_adapters = s["nic_adapters"] or []
        psus = s["psus"] or []

        sys_sn = system.get("SerialNumber") or ""
        mb_sn = chassis.get("SerialNumber") or ""

        # Supermicro: chassis SN = system serial
        if "supermicro" in vendor:
            ch_sn = sys_sn
        else:
            ch_sn = ""

        if ch_sn:
            cols["Chassis SN:"].append(ch_sn)
        if mb_sn:
            cols["MB SN:"].append(mb_sn)
        if sys_sn:
            cols["Blade SN:"].append(sys_sn)
            cols["Server SN:"].append(sys_sn)

        # CPU serials (if any)
        for p in procs:
            sn = p.get("SerialNumber")
            if not sn:
                oem = p.get("Oem") or {}
                for v in oem.values():
                    if isinstance(v, dict):
                        for k2, v2 in v.items():
                            if isinstance(v2, str) and "serial" in k2.lower():
                                sn = v2
                                break
                    if sn:
                        break
            if sn:
                cols["CPU SN:"].append(sn)

        # DIMM serials
        for d in dimms:
            sn = d.get("SerialNumber")
            if sn:
                cols["Memory SN:"].append(sn)

        # NIC card serials: only true serial numbers (no MAC fallback)
        nic_sn_list = []
        for a in nic_adapters:
            sn = a.get("SerialNumber")
            if not sn:
                oem = a.get("Oem") or {}
                for v in oem.values():
                    if isinstance(v, dict):
                        for k2, v2 in v.items():
                            if isinstance(v2, str) and "serial" in k2.lower():
                                sn = v2
                                break
                    if sn:
                        break
            if sn and sn not in nic_sn_list:
                nic_sn_list.append(sn)

        for sn in nic_sn_list:
            cols["NIC card SN:"].append(sn)

        # NIC MAC addresses from EthernetInterfaces
        nic_mac_list = []
        for n in nic_ports:
            mac = n.get("MACAddress")
            if mac and mac not in nic_mac_list:
                nic_mac_list.append(mac)

        for mac in nic_mac_list:
            cols["NIC MAC:"].append(mac)

        # Drive serials
        for drv in drives:
            sn = drv.get("_BetterSerial") or drv.get("SerialNumber")
            if sn:
                cols["Storage SN:"].append(sn)

        # PSU serials – duplicates allowed (multi-node shares PSU)
        for p in psus:
            sn = p.get("SerialNumber") or p.get("Model")
            if sn:
                cols["PSU SN:"].append(sn)

    return cols

def write_serials_excel(path_xlsx, servers):
    cols = build_serial_columns(servers)
    wb = Workbook()
    ws = wb.active
    headers = [
        "Chassis SN:",
        "MB SN:",
        "Blade SN:",
        "CPU SN:",
        "Memory SN:",
        "NIC card SN:",
        "NIC MAC:",
        "Storage SN:",
        "PSU SN:",
        "Server SN:",
    ]
    ws.append(headers)
    maxlen = max((len(cols[h]) for h in headers), default=0)
    for i in range(maxlen):
        row = []
        for h in headers:
            row.append(cols[h][i] if i < len(cols[h]) else "")
        ws.append(row)
    wb.save(path_xlsx)

# --------------------------------------------------------------------
# Config comparison (group PDF)
# --------------------------------------------------------------------
def fingerprint_for_group(s):
    procs = tuple(sorted([p.get("Model") for p in s["procs"] if p.get("Model")]))
    dimm_pn = tuple(sorted([(d.get("PartNumber") or "") for d in s["dimms"]]))
    mb_pn = (s["system"].get("PartNumber") or s["chassis"].get("PartNumber") or "")
    drive_models = tuple(sorted([(d.get("Model") or "") for d in s["drives"]]))
    psu_models = tuple(sorted([(p.get("Model") or "") for p in s["psus"]]))
    return (procs, dimm_pn, mb_pn, drive_models, psu_models)

def configs_all_equal(servers):
    if not servers:
        return True
    sig0 = fingerprint_for_group(servers[0])
    for s in servers[1:]:
        if fingerprint_for_group(s) != sig0:
            return False
    return True

def group_extras_differences(servers):
    notes = []
    base = servers[0]

    def counts_cpu(ss): return Counter([p.get("Model") for p in ss["procs"] if p.get("Model")])
    def counts_dimm(ss): return Counter([(d.get("PartNumber") or "") for d in ss["dimms"]])
    def counts_drive(ss): return Counter([(d.get("Model") or "") for d in ss["drives"]])
    def counts_psu(ss):  return Counter([(p.get("Model") or "") for p in ss["psus"]])

    for s in servers[1:]:
        diff = []
        if counts_cpu(s) != counts_cpu(base):
            diff.append(f"CPU:{dict(counts_cpu(s))}")
        if counts_dimm(s) != counts_dimm(base):
            diff.append(f"MEM:{dict(counts_dimm(s))}")
        if counts_drive(s) != counts_drive(base):
            diff.append(f"DRV:{dict(counts_drive(s))}")
        if counts_psu(s) != counts_psu(base):
            diff.append(f"PSU:{dict(counts_psu(s))}")
        if diff:
            notes.append(f"{s['ip']} -> " + "; ".join(diff))
    return " | ".join(notes)

# --------------------------------------------------------------------
# High level description for PDF
# --------------------------------------------------------------------
def describe_for_pdf(server):
    system = server["system"] or {}
    chassis = server["chassis"] or {}
    procs = server["procs"] or []
    dimms = server["dimms"] or []
    drives = server["drives"] or []
    nic_adapters = server.get("nic_adapters") or []
    psus = server["psus"] or []

    chassis_model = chassis.get("Model") or system.get("Model") or ""
    board_pn = system.get("PartNumber") or chassis.get("PartNumber") or ""
    system_serial = system.get("SerialNumber") or chassis.get("SerialNumber") or ""

    # If chassis model and board PN are identical, prefer MB in the field
    motherboard = board_pn or chassis_model

    cpu_txt = summarize_cpu(procs)
    mem_txt = summarize_memory(dimms)
    drv_txt = summarize_drives(drives)
    raid_txt = ""  # unified sürümde RAID hw toplanmıyor (istersek ekleriz)
    nic_txt = summarize_network(nic_adapters)
    psu_txt = summarize_psus(psus)
    bios_txt = (server.get("bios") or "").strip()
    bmc_txt = (server.get("bmcver") or "").strip()

    return {
        "CHASSIS_PLATFORM": chassis_model,
        "MOTHERBOARD": motherboard,
        "SYSTEM_SERIALS": system_serial,
        "CPU": cpu_txt,
        "MEMORY": mem_txt,
        "DRIVES": drv_txt,
        "RAID_HW": raid_txt,
        "NETWORK_HW": nic_txt,
        "PSU": psu_txt,
        "BIOS": bios_txt,
        "BMC_ILO": bmc_txt,
        "EXTRAS": "",
    }

def fill_pdf_single(server, out_pdf, quantity=1):
    vals = describe_for_pdf(server)
    vals["QUANTITY"] = str(quantity)
    write_pdf_from_values(server["ip"], vals, out_pdf)

def fill_pdf_group(servers, out_pdf):
    if configs_all_equal(servers):
        fill_pdf_single(servers[0], out_pdf, quantity=len(servers))
        return
    vals = describe_for_pdf(servers[0])
    vals["QUANTITY"] = str(len(servers))
    vals["EXTRAS"] = group_extras_differences(servers)
    write_pdf_from_values(servers[0]["ip"], vals, out_pdf)

# --------------------------------------------------------------------
# Redfish data collection wrapper
# --------------------------------------------------------------------
def collect_one(ip, user, password, vendor_key):
    auth = (user, password)
    sys_path, ch_path, mgr_path = discover_paths(ip, auth, vendor_key)

    system = get_system(ip, auth, sys_path)
    chassis = get_chassis(ip, auth, ch_path)
    procs = get_processors(ip, auth, sys_path)
    dimms = get_dimms(ip, auth, sys_path)
    drives = get_drives(ip, auth, sys_path, ch_path)
    nic_ports = get_nic_ports(ip, auth, sys_path, mgr_path)
    nic_adapters = get_nic_adapters(ip, auth, sys_path, ch_path)
    psus = get_psus(ip, auth, ch_path)
    bios = get_bios_version(system) or ""
    bmcv = get_bmc_version(ip, auth, mgr_path) or ""

    data = {
        "ip": ip,
        "vendor": (system.get("Manufacturer") or vendor_key or "").lower(),
        "system": system,
        "chassis": chassis,
        "procs": procs,
        "dimms": dimms,
        "drives": drives,
        "nic_ports": nic_ports,
        "nic_adapters": nic_adapters,
        "psus": psus,
        "bios": bios,
        "bmcver": bmcv,
    }
    return data

# --------------------------------------------------------------------
# CSV & CLI
# --------------------------------------------------------------------
def read_task_csv(path_csv):
    items = []
    if not os.path.exists(path_csv):
        return items
    with open(path_csv, encoding="utf-8") as f:
        rd = csv.DictReader(f)
        for r in rd:
            ip = (r.get("IP") or "").strip()
            user = (r.get("User") or "").strip()
            pw = (r.get("Password") or "").strip()
            if ip and user:
                items.append((ip, user, pw))
    return items

def ask_vendor():
    print("Vendor seçin:")
    print("  1 - Dell iDRAC")
    print("  2 - HPE iLO")
    print("  3 - Supermicro BMC")
    ch = input("Seçim (1/2/3): ").strip()
    if ch == "1":
        return "dell"
    if ch == "2":
        return "hpe"
    if ch == "3":
        return "supermicro"
    return "dell"

def ask_folder_name():
    name = input("Çıktı klasör adı (ör. SO# 32252): ").strip()
    if not name:
        name = datetime.now().strftime("RUN_%Y%m%d_%H%M%S")
    return name

def main():
    print("=== Unified Server Configurator (CLI) ===")
    vendor_key = ask_vendor()

    csv_path = os.path.join(BASE_DIR, "task.csv")
    rows = read_task_csv(csv_path)
    if not rows:
        print(f"task.csv boş ya da bulunamadı: {csv_path}")
        input("Kapatmak için Enter...")
        return

    out_root = DEFAULT_OUTPUT_DIRS.get(vendor_key, DEFAULT_OUTPUT_DIRS["hpe"])
    folder_name = ask_folder_name()
    out_dir = os.path.join(out_root, folder_name)
    os.makedirs(out_dir, exist_ok=True)

    servers = []
    for ip, user, pw in rows:
        print(f"[{ip}] Redfish verileri alınıyor...")
        try:
            sdata = collect_one(ip, user, pw, vendor_key)
            if not sdata.get("vendor"):
                sdata["vendor"] = vendor_key
            servers.append(sdata)
        except Exception as e:
            print(f"[{ip}] HATA: {e}")

    if not servers:
        print("Hiç sunucu okunamadı.")
        input("Kapatmak için Enter...")
        return

    # Individual PDFs
    for s in servers:
        pdf_name = f"Configurator_{s['ip'].replace(':', '_')}.pdf"
        out_pdf = os.path.join(out_dir, pdf_name)
        fill_pdf_single(s, out_pdf, quantity=1)

    # Group summary PDF
    group_pdf = os.path.join(out_dir, "Configurator_GROUP.pdf")
    fill_pdf_group(servers, group_pdf)

    # Excel with serials
    xlsx_path = os.path.join(out_dir, f"{folder_name}_serials.xlsx")
    write_serials_excel(xlsx_path, servers)

    print("\nÇıktılar hazır:")
    print(f"  {out_dir}")
    for fn in sorted(os.listdir(out_dir)):
        print("  -", fn)
    input("\nBitti. Kapatmak için Enter...")

if __name__ == "__main__":
    main()

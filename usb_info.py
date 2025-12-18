from __future__ import annotations

import json
import re
import subprocess
from typing import Any, Dict, List, Optional


_VID_PID_RE = re.compile(r"VID_([0-9A-Fa-f]{4}).*PID_([0-9A-Fa-f]{4})")
_SERIAL_FROM_PNP_RE = re.compile(r"^USB\\[^\\]+\\([^\\]+)$", re.IGNORECASE)


def _run_powershell_json(ps_script: str) -> Any:
    prefix = r"""
$ErrorActionPreference = 'Stop';
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
$OutputEncoding = [System.Text.Encoding]::UTF8;
"""
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", prefix + "\n" + ps_script]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode != 0:
        raise RuntimeError(f"PowerShell 执行失败：{p.stderr.strip() or p.stdout.strip()}")
    out = p.stdout.strip()
    if not out:
        return []
    return json.loads(out)


def _parse_vid_pid(pnp_device_id: Optional[str]) -> Dict[str, Optional[str]]:
    if not pnp_device_id:
        return {"vendor_id": None, "product_id": None}
    m = _VID_PID_RE.search(pnp_device_id)
    if not m:
        return {"vendor_id": None, "product_id": None}
    return {"vendor_id": f"0x{m.group(1).lower()}", "product_id": f"0x{m.group(2).lower()}"}


def _parse_serial(pnp_device_id: Optional[str]) -> Optional[str]:
    if not pnp_device_id:
        return None
    m = _SERIAL_FROM_PNP_RE.match(pnp_device_id)
    return m.group(1) if m else None


def list_usb_devices(only_storage: bool = True) -> List[Dict[str, Any]]:
    """
    Windows：WMI 枚举 USB 设备。

    only_storage=True：仅显示 USB 存储设备（Service=USBSTOR），更贴合“U盘检测”实验。
    only_storage=False：显示全部 USB 设备（用于扩展功能/调试）。
    """
    ps = r"""
$rows = Get-CimInstance Win32_PnPEntity |
  Where-Object { $_.PNPDeviceID -like 'USB*' } |
  Select-Object Name, Manufacturer, PNPDeviceID, Service;

$rows | ConvertTo-Json -Depth 4
"""
    data = _run_powershell_json(ps)
    rows = [data] if isinstance(data, dict) else (data or [])

    devices: List[Dict[str, Any]] = []
    for r in rows:
        service = (r.get("Service") or "").upper()
        if only_storage and service != "USBSTOR":
            continue

        name = r.get("Name")
        manufacturer = r.get("Manufacturer")
        pnp = r.get("PNPDeviceID")

        vidpid = _parse_vid_pid(pnp)
        serial = _parse_serial(pnp)

        devices.append(
            {
                "vendor_id": vidpid["vendor_id"],
                "product_id": vidpid["product_id"],
                "manufacturer": manufacturer,
                "product": name,
                "serial_number": serial,
                "usb_version_bcd": None,
                "bus": None,
                "address": None,
                "pnp_device_id": pnp,
                "service": r.get("Service"),
            }
        )

    return devices
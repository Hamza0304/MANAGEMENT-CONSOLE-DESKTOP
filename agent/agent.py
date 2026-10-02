
import json
import logging
import os
import platform
import socket
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Any

import psutil
import requests

AGENT_VERSION = "1.1.0"
AGENT_INTERVAL = int(os.environ.get("AGENT_INTERVAL", "300"))
REQUEST_TIMEOUT = 30


def _default_log_dir() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("PROGRAMDATA", Path.home()))
    else:
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    return base / "ManagementConsoleAgent" / "logs"


def configure_logging() -> logging.Logger:
    log_dir = _default_log_dir()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "agent.log"

    logger = logging.getLogger("management_console_agent")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


logger = configure_logging()


def resolve_config_path(base_dir: Path | None = None, override_path: str | None = None) -> Path:
    if override_path:
        return Path(override_path).expanduser()

    explicit = os.environ.get("MANAGEMENT_CONSOLE_AGENT_CONFIG", "").strip()
    if explicit:
        return Path(explicit).expanduser()

    if base_dir is None:
        if getattr(sys, "frozen", False):
            base_dir = Path(sys.executable).resolve().parent
        else:
            base_dir = Path(__file__).resolve().parent

    candidates = [
        base_dir / "config.json",
        Path.cwd() / "config.json",
        Path(os.environ.get("PROGRAMDATA", "")) / "ManagementConsoleAgent" / "config.json",
        Path(os.environ.get("LOCALAPPDATA", "")) / "ManagementConsoleAgent" / "config.json",
        Path(__file__).resolve().parent.parent / "config.json",
    ]

    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate

    return candidates[0]


CONFIG_FILE = resolve_config_path()


def safe_text(value: Any, fallback: str = "Unknown") -> str:
    if value is None:
        return fallback
    value = str(value).strip()
    return value if value else fallback


def normalize_network_values(values: Any, fallback: str = "Unknown") -> str:
    if isinstance(values, (list, tuple, set)):
        cleaned = [safe_text(v, fallback) for v in values if v not in (None, "", "Unknown")]
        if cleaned:
            return ", ".join(cleaned)
        return fallback
    if values is None or values == "":
        return fallback
    return safe_text(values, fallback)


def normalize_inventory_record(record: dict[str, Any] | None) -> dict[str, Any]:
    record = record or {}

    computer_name = safe_text(record.get("computer_name") or socket.gethostname(), "Unknown")
    manufacturer = safe_text(record.get("manufacturer"), "Unknown")
    model = safe_text(record.get("model"), "Unknown")
    serial_number = safe_text(record.get("serial_number"), "Unknown")
    windows = safe_text(record.get("windows") or record.get("operating_system"), "Unknown")
    os_version = safe_text(record.get("os_version") or record.get("windows_version"), "Unknown")
    architecture = safe_text(record.get("architecture"), platform.machine() or "Unknown")

    cpu = record.get("cpu") or {}
    ram = record.get("ram") or {}
    disk = record.get("disk") or {}

    normalized = {
        "computer_name": computer_name,
        "manufacturer": manufacturer,
        "model": model,
        "serial_number": serial_number,
        "windows": windows,
        "os_version": os_version,
        "architecture": architecture,
        "cpu": {
            "name": safe_text(cpu.get("name"), "Unknown"),
            "cores": cpu.get("cores"),
            "logical_processors": cpu.get("logical_processors"),
        },
        "ram": {
            "total_gb": ram.get("total_gb"),
            "available_gb": ram.get("available_gb"),
            "usage_percent": ram.get("usage_percent"),
        },
        "disk": {
            "total_gb": disk.get("total_gb"),
            "free_gb": disk.get("free_gb"),
            "usage_percent": disk.get("usage_percent"),
        },
        "ip_address": normalize_network_values(record.get("ip_address") or record.get("ip_addresses"), "Unknown"),
        "mac_address": normalize_network_values(record.get("mac_address") or record.get("mac_addresses"), "Unknown"),
        "installed_software": collect_installed_software_entries(record.get("installed_software") or []),
    }

    if record.get("ip_addresses") and isinstance(record.get("ip_addresses"), list):
        normalized["ip_addresses"] = [safe_text(item, "Unknown") for item in record["ip_addresses"] if safe_text(item, "Unknown") != "Unknown"]
    if record.get("mac_addresses") and isinstance(record.get("mac_addresses"), list):
        normalized["mac_addresses"] = [safe_text(item, "Unknown") for item in record["mac_addresses"] if safe_text(item, "Unknown") != "Unknown"]

    return normalized


def run_windows_command(args: list[str], timeout: int = 20000):
    kwargs: dict[str, Any] = {
        "capture_output": True,
        "text": True,
        "check": False,
        "timeout": timeout,
    }
    if os.name == "nt":
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.run(args, **kwargs)


def get_cpu_info() -> dict[str, Any]:
    info = {
        "name": platform.processor() or "Unknown",
        "cores": psutil.cpu_count(logical=False),
        "logical_processors": psutil.cpu_count(logical=True),
    }
    try:
        result = run_windows_command([
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            "(Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name).ToString()",
        ], timeout=15000)
        if result.returncode == 0 and result.stdout.strip():
            info["name"] = result.stdout.strip()
    except Exception:
        logger.debug("CPU name query failed", exc_info=True)
    return info


def get_ram_info() -> dict[str, Any]:
    ram = psutil.virtual_memory()
    return {
        "total_gb": round(ram.total / (1024 ** 3), 2),
        "available_gb": round(ram.available / (1024 ** 3), 2),
        "usage_percent": ram.percent,
    }


def get_disk_info() -> dict[str, Any]:
    disk = psutil.disk_usage("C:\\")
    return {
        "total_gb": round(disk.total / (1024 ** 3), 2),
        "free_gb": round(disk.free / (1024 ** 3), 2),
        "usage_percent": disk.percent,
    }


def get_windows_system_details() -> dict[str, str]:
    details: dict[str, str] = {
        "manufacturer": "Unknown",
        "model": "Unknown",
        "serial_number": "Unknown",
        "windows": "Unknown",
        "os_version": "Unknown",
        "architecture": platform.machine() or "Unknown",
    }

    script = r"""
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    [pscustomobject]@{
      manufacturer = if ($cs) { [string]$cs.Manufacturer } else { 'Unknown' }
      model = if ($cs) { [string]$cs.Model } else { 'Unknown' }
      serial_number = if ($bios) { [string]$bios.SerialNumber } else { 'Unknown' }
      windows = if ($os) { [string]$os.Caption } else { 'Unknown' }
      os_version = if ($os) { [string]$os.Version } else { 'Unknown' }
      architecture = if ($os) { [string]$os.OSArchitecture } else { 'Unknown' }
    } | ConvertTo-Json -Compress
    """

    try:
        result = run_windows_command([
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ], timeout=20000)
        if result.returncode == 0 and result.stdout.strip():
            payload = json.loads(result.stdout)
            for key, value in payload.items():
                if value not in (None, ""):
                    details[key] = str(value).strip()
    except Exception:
        logger.debug("Windows system detail query failed", exc_info=True)

    return details


def get_ip_addresses() -> list[str]:
    addresses: list[str] = []
    try:
        result = run_windows_command([
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            "(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter \"IPEnabled = True\" -ErrorAction SilentlyContinue | ForEach-Object { $_.IPAddress } | Where-Object { $_ -and $_ -notmatch '::' } | Select-Object -Unique) | ConvertTo-Json -Compress",
        ], timeout=20000)
        if result.returncode == 0 and result.stdout.strip():
            payload = json.loads(result.stdout)
            if isinstance(payload, list):
                addresses = [str(value).strip() for value in payload if str(value).strip() and not str(value).startswith("127.")]
            elif isinstance(payload, str):
                addresses = [payload.strip()]
    except Exception:
        logger.debug("IP address query failed", exc_info=True)

    if not addresses:
        for addr in (psutil.net_if_addrs() or {}).values():
            for item in addr:
                if item.family == socket.AF_INET and item.address and not item.address.startswith("127."):
                    addresses.append(item.address)
    return list(dict.fromkeys(addresses))


def get_mac_addresses() -> list[str]:
    addresses: list[str] = []
    try:
        result = run_windows_command([
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            "(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter \"IPEnabled = True\" -ErrorAction SilentlyContinue | Where-Object { $_.MacAddress } | Select-Object -ExpandProperty MacAddress -Unique) | ConvertTo-Json -Compress",
        ], timeout=20000)
        if result.returncode == 0 and result.stdout.strip():
            payload = json.loads(result.stdout)
            if isinstance(payload, list):
                addresses = [str(value).strip() for value in payload if str(value).strip()]
            elif isinstance(payload, str):
                addresses = [payload.strip()]
    except Exception:
        logger.debug("MAC address query failed", exc_info=True)

    if not addresses:
        try:
            mac = uuid.getnode()
            if mac:
                addresses = [":".join(f"{(mac >> i) & 0xff:02X}" for i in range(40, -1, -8))]
        except Exception:
            logger.debug("Fallback MAC generation failed", exc_info=True)
    return list(dict.fromkeys(addresses))


def collect_installed_software_entries(entries: Any) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()

    if not isinstance(entries, list):
        return items

    for entry in entries:
        if not isinstance(entry, dict):
            continue

        name = safe_text(entry.get("name") or entry.get("DisplayName") or entry.get("display_name"), "Unknown")
        version = safe_text(entry.get("version") or entry.get("DisplayVersion") or entry.get("display_version"), "Unknown")
        publisher = safe_text(entry.get("publisher") or entry.get("Publisher") or entry.get("manufacturer"), "Unknown")
        install_date = safe_text(entry.get("install_date") or entry.get("InstallDate"), "Unknown")
        architecture = safe_text(entry.get("architecture"), "Unknown")

        if name == "Unknown":
            continue

        key = (name.lower(), version.lower(), publisher.lower())
        if key in seen:
            continue
        seen.add(key)

        items.append({
            "name": name,
            "version": version,
            "publisher": publisher,
            "install_date": install_date if install_date != "Unknown" else None,
            "architecture": architecture,
        })

    return items


def get_installed_software() -> list[dict[str, str]]:
    script = r"""
    $paths = @(
      'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
      'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
      'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $results = foreach ($path in $paths) {
      try {
        Get-ItemProperty $path -ErrorAction Stop
      } catch {
        $null
      }
    }

    $rows = foreach ($item in $results) {
      $name = $item.DisplayName
      if ([string]::IsNullOrWhiteSpace($name)) { continue }
      $publisher = if ($item.Publisher) { [string]$item.Publisher } else { 'Unknown' }
      $version = if ($item.DisplayVersion) { [string]$item.DisplayVersion } else { 'Unknown' }
      $installDate = if ($item.InstallDate) { [string]$item.InstallDate } else { $null }
      $architecture = if ($item.PSPath -match 'Wow6432Node') { '32-bit' } else { '64-bit' }
      [pscustomobject]@{
        name = [string]$name
        publisher = [string]$publisher
        version = [string]$version
        install_date = $installDate
        architecture = $architecture
      }
    }

    $rows | Sort-Object name, version | ConvertTo-Json -Depth 5 -Compress
    """

    try:
        result = run_windows_command([
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ], timeout=25000)
        if result.returncode != 0:
            logger.warning("Installed software query failed with return code %s: %s", result.returncode, result.stderr.strip())
            return []

        payload = result.stdout.strip()
        if not payload:
            return []

        parsed = json.loads(payload)
        if isinstance(parsed, list):
            return collect_installed_software_entries(parsed)
        return collect_installed_software_entries([parsed])
    except Exception:
        logger.exception("Unable to collect installed software")
        return []


def get_system_info() -> dict[str, Any]:
    def collect(name: str, collector, default):
        try:
            value = collector()
            logger.info("Collected %s: %s", name, value)
            return value
        except Exception:
            logger.exception("Unable to collect %s", name)
            return default

    inventory = {
        "computer_name": collect("computer name", socket.gethostname, "Unknown"),
        "manufacturer": "Unknown",
        "model": "Unknown",
        "serial_number": "Unknown",
        "cpu": collect("CPU", get_cpu_info, {}),
        "ram": collect("RAM", get_ram_info, {}),
        "disk": collect("storage", get_disk_info, {}),
        "installed_software": collect("installed software", get_installed_software, []),
    }

    windows_details = collect("Windows inventory", get_windows_system_details, {})
    ip_addresses = collect("IP addresses", get_ip_addresses, [])
    mac_addresses = collect("MAC addresses", get_mac_addresses, [])

    inventory["manufacturer"] = safe_text(windows_details.get("manufacturer"), "Unknown")
    inventory["model"] = safe_text(windows_details.get("model"), "Unknown")
    inventory["serial_number"] = safe_text(windows_details.get("serial_number"), "Unknown")
    inventory["windows"] = safe_text(windows_details.get("windows"), platform.platform())
    inventory["os_version"] = safe_text(windows_details.get("os_version"), platform.version())
    inventory["architecture"] = safe_text(windows_details.get("architecture"), platform.machine() or "Unknown")
    inventory["ip_address"] = normalize_network_values(ip_addresses, "Unknown")
    inventory["ip_addresses"] = [safe_text(value, "Unknown") for value in ip_addresses if safe_text(value, "Unknown") != "Unknown"]
    inventory["mac_address"] = normalize_network_values(mac_addresses, "Unknown")
    inventory["mac_addresses"] = [safe_text(value, "Unknown") for value in mac_addresses if safe_text(value, "Unknown") != "Unknown"]
    inventory["installed_software"] = collect_installed_software_entries(inventory.get("installed_software") or [])

    logger.info("Inventory fields collected: %s", len(inventory))
    return normalize_inventory_record(inventory)


def load_config() -> dict[str, str]:
    config_path = resolve_config_path()
    logger.info("Configuration source: %s", config_path)

    if not config_path.exists():
        config_path.parent.mkdir(parents=True, exist_ok=True)
        default = {
            "server_url": os.environ.get("MANAGEMENT_CONSOLE_SERVER_URL", "http://127.0.0.1:5000"),
            "api_key": "",
        }
        config_path.write_text(json.dumps(default, indent=2), encoding="utf-8")
        logger.warning("Configuration file was missing; created a default config at %s", config_path)

    try:
        with config_path.open("r", encoding="utf-8") as file:
            config = json.load(file)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON configuration: {config_path}") from exc

    server_url = str(config.get("server_url") or os.environ.get("MANAGEMENT_CONSOLE_SERVER_URL", "")).strip()
    api_key = str(config.get("api_key") or os.environ.get("MANAGEMENT_CONSOLE_API_KEY", "")).strip()

    if not server_url:
        raise ValueError("server_url is missing from config.json")

    server_url = server_url.rstrip("/")
    if not server_url.endswith("/api/agent"):
        server_url += "/api/agent"

    logger.info("Agent installation directory: %s", str(Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent))
    logger.info("Agent version: %s", AGENT_VERSION)
    logger.info("Server endpoint: %s", server_url)

    return {
        "server_url": server_url,
        "api_key": api_key,
        "config_file": str(config_path),
    }


def send_system_info(information: dict[str, Any], config: dict[str, str]) -> bool:
    headers: dict[str, str] = {}
    if config.get("api_key"):
        headers["Authorization"] = f"Bearer {config['api_key']}"

    logger.info("API request started for %s", config["server_url"])
    try:
        response = requests.post(
            config["server_url"],
            json=information,
            headers=headers,
            timeout=REQUEST_TIMEOUT,
        )
    except requests.RequestException:
        logger.exception("API unavailable while sending inventory")
        return False

    logger.info("API response status: %s", response.status_code)
    if response.ok:
        logger.info("Successful synchronization with server.")
        return True

    try:
        payload = response.json()
        logger.warning("API rejected the update with status %s: %s", response.status_code, payload)
    except ValueError:
        logger.warning("API rejected the update with status %s: %s", response.status_code, response.text)
    return False


def run_agent() -> None:
    try:
        config = load_config()
    except Exception:
        logger.exception("Unable to load agent configuration")
        return

    logger.info("Agent startup complete. Update interval: %s seconds", AGENT_INTERVAL)

    try:
        while True:
            try:
                logger.info("Data collection started.")
                information = get_system_info()
                logger.info("Data collection completed with %s fields.", len(information))
                send_system_info(information, config)
            except Exception:
                logger.exception("Unexpected error during agent update")

            logger.info("Next update in %s seconds", AGENT_INTERVAL)
            try:
                time.sleep(AGENT_INTERVAL)
            except KeyboardInterrupt:
                break
    except KeyboardInterrupt:
        logger.info("Agent shutdown requested by user.")
    finally:
        logger.info("Agent stopped.")


if __name__ == "__main__":
    run_agent()


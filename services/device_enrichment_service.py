from __future__ import annotations

import json
import logging
import os
import re
import socket
import subprocess
from typing import Any


logger = logging.getLogger(__name__)


# ============================================================
# BASIC HELPERS
# ============================================================

def normalize_mac_address(value: Any) -> str | None:
    """Normalize a MAC address into AA:BB:CC:DD:EE:FF format."""
    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    hex_chars = re.sub(r"[^0-9A-Fa-f]", "", text)

    if len(hex_chars) != 12:
        return text.upper()

    return ":".join(
        hex_chars[index:index + 2]
        for index in range(0, 12, 2)
    ).upper()


def normalize_host_name(value: Any) -> str | None:
    """Normalize a hostname."""
    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    return text


def safe_text(value: Any, default: str = "") -> str:
    """Convert a value safely to text."""
    if value is None:
        return default

    try:
        return str(value).strip()
    except Exception:
        return default


def run_windows_command(
    command: list[str],
    timeout: int = 30,
) -> tuple[int, str, str]:
    """
    Execute a Windows command safely.

    Returns:
        (return_code, stdout, stderr)
    """
    try:
        process_options: dict[str, Any] = {
            "capture_output": True,
            "text": True,
            "timeout": timeout,
            "encoding": "utf-8",
            "errors": "replace",
            "shell": False,
        }
        if os.name == "nt":
            startup_info = subprocess.STARTUPINFO()
            startup_info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startup_info.wShowWindow = subprocess.SW_HIDE
            process_options["startupinfo"] = startup_info
            process_options["creationflags"] = subprocess.CREATE_NO_WINDOW

        process = subprocess.run(command, **process_options)

        return (
            process.returncode,
            process.stdout or "",
            process.stderr or "",
        )

    except subprocess.TimeoutExpired as exc:
        stdout = exc.stdout or ""
        stderr = exc.stderr or ""

        if isinstance(stdout, bytes):
            stdout = stdout.decode("utf-8", errors="replace")

        if isinstance(stderr, bytes):
            stderr = stderr.decode("utf-8", errors="replace")

        return (
            -1,
            stdout,
            f"Command timed out: {stderr}",
        )

    except Exception as exc:
        return (
            -1,
            "",
            str(exc),
        )


def extract_local_ip_addresses() -> list[str]:
    """Return IPv4 addresses assigned to active local adapters."""
    addresses: list[str] = []

    # On Windows, get addresses from adapters that are actually up. This
    # avoids treating stale addresses on disconnected interfaces as local.
    if os.name == "nt":
        command = [
            "powershell.exe", "-NoProfile", "-NonInteractive",
            "-ExecutionPolicy", "Bypass", "-Command",
            "@(Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq 'Up' } "
            "| ForEach-Object { $_.IPv4Address.IPAddress } | Where-Object { $_ }) "
            "| ConvertTo-Json -Compress",
        ]
        code, stdout, _ = run_windows_command(command, timeout=10)
        if code == 0:
            parsed = parse_powershell_json(stdout)
            if isinstance(parsed, str):
                parsed = [parsed]
            if isinstance(parsed, list):
                for ip in parsed:
                    if ip and ip not in addresses and not str(ip).startswith("127."):
                        addresses.append(str(ip))

    if not addresses:
        try:
            hostname = socket.gethostname()

            for item in socket.getaddrinfo(
                hostname,
                None,
                socket.AF_INET,
                socket.SOCK_STREAM,
            ):
                ip = item[4][0]

                if ip and not ip.startswith("127.") and ip not in addresses:
                    addresses.append(ip)

        except Exception as exc:
            logger.debug(
                "Could not resolve local hostname addresses: %s",
                exc,
            )

    try:
        with socket.socket(
            socket.AF_INET,
            socket.SOCK_DGRAM,
        ) as sock:
            sock.connect(("8.8.8.8", 80))
            local_ip = sock.getsockname()[0]

            if local_ip and local_ip not in addresses:
                addresses.append(local_ip)

    except Exception as exc:
        logger.debug(
            "Could not determine primary local IP: %s",
            exc,
        )

    return addresses


def query_local_windows_enrichment(
    ip_address: str,
    mac_address: str | None = None,
    hostname: str | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    """Collect local Windows CIM data without creating a DCOM session."""
    script = r"""
$ErrorActionPreference = 'Stop'
try {
    $computer = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $processors = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
    $bios = Get-CimInstance Win32_BIOS -ErrorAction Stop
    $configs = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' -ErrorAction Stop)
    $adapters = @(Get-CimInstance Win32_NetworkAdapter -ErrorAction SilentlyContinue)
    $disks = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType = 3' -ErrorAction SilentlyContinue)
    $net = @($configs | ForEach-Object {
        $cfg = $_
        $adapter = $adapters | Where-Object { $_.MACAddress -eq $cfg.MACAddress } | Select-Object -First 1
        [PSCustomObject]@{ name=$cfg.Description; mac_address=$cfg.MACAddress; ip_addresses=@($cfg.IPAddress); gateway=@($cfg.DefaultIPGateway); dns_servers=@($cfg.DNSServerSearchOrder); dhcp_enabled=$cfg.DHCPEnabled; vendor=$adapter.Manufacturer; speed=$adapter.Speed }
    })
    $diskData = @($disks | ForEach-Object {
        [PSCustomObject]@{ device_id=$_.DeviceID; volume_name=$_.VolumeName; size_gb=[math]::Round(([double]$_.Size / 1GB),2); free_gb=[math]::Round(([double]$_.FreeSpace / 1GB),2); used_gb=[math]::Round((([double]$_.Size - [double]$_.FreeSpace) / 1GB),2) }
    })
    $osData = [PSCustomObject]@{
        success=$true; access_status='Success'; ip_address='__TARGET_IP__'; mac_address='__TARGET_MAC__'; hostname=$computer.Name
        device_type=$(if ($computer.PCSystemType -eq 2 -or $computer.PCSystemType -eq 8) { 'Laptop' } else { 'Desktop' })
        manufacturer=$computer.Manufacturer; model=$computer.Model; serial_number=$bios.SerialNumber
        domain=$(if ($computer.PartOfDomain) { $computer.Domain } else { '' })
        workgroup=$(if (-not $computer.PartOfDomain) { $computer.Domain } else { '' })
        logged_in_user=$computer.UserName; operating_system=$os.Caption; windows_version=$os.Version; build_number=$os.BuildNumber
        architecture=$os.OSArchitecture; registered_user=$os.RegisteredUser; os_install_date=$(if ($os.InstallDate) { $os.InstallDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' })
        ram_gb=[math]::Round(([double]$computer.TotalPhysicalMemory / 1GB),2)
        cpu=[PSCustomObject]@{ name=$processors[0].Name; manufacturer=$processors[0].Manufacturer; cores=$processors[0].NumberOfCores; logical_processors=$processors[0].NumberOfLogicalProcessors; max_clock_mhz=$processors[0].MaxClockSpeed }
        bios=[PSCustomObject]@{ manufacturer=$bios.Manufacturer; version=$bios.SMBIOSBIOSVersion; serial_number=$bios.SerialNumber; release_date=$(if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString('yyyy-MM-dd') } else { '' }) }
        network=$net; ip_addresses=@($configs | ForEach-Object { $_.IPAddress } | Where-Object { $_ -notmatch ':' } | Select-Object -Unique)
        disks=$diskData; installed_software=@(); installed_software_count=0; enrichment_status='Complete'
    }
    $osData | ConvertTo-Json -Depth 10 -Compress
} catch {
    [PSCustomObject]@{ success=$false; access_status='Error'; ip_address='__TARGET_IP__'; mac_address='__TARGET_MAC__'; hostname='__TARGET_HOSTNAME__'; reason=$_.Exception.Message } | ConvertTo-Json -Compress
}
"""
    script = (script.replace("__TARGET_IP__", safe_text(ip_address))
              .replace("__TARGET_MAC__", safe_text(mac_address))
              .replace("__TARGET_HOSTNAME__", safe_text(hostname)))
    command = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script]
    code, stdout, stderr = run_windows_command(command, timeout=timeout)
    parsed = parse_powershell_json(stdout)
    if code != 0 or not isinstance(parsed, dict) or not parsed.get("success"):
        reason = safe_text(parsed.get("reason") if isinstance(parsed, dict) else "") or _enrichment_reason_from_output(stdout, stderr)
        return build_enrichment_result(success=False, ip_address=ip_address, mac_address=mac_address, hostname=hostname, data=parsed if isinstance(parsed, dict) else {}, reason=reason, access_status=safe_text(parsed.get("access_status")) if isinstance(parsed, dict) else "Error")

    parsed["installed_software"] = get_local_installed_software()
    parsed["installed_software_count"] = len(parsed["installed_software"])
    target_mac = mac_address
    for adapter in parsed.get("network", []):
        if ip_address in adapter.get("ip_addresses", []):
            target_mac = adapter.get("mac_address") or target_mac
            break
    return build_enrichment_result(success=True, ip_address=ip_address, mac_address=target_mac, hostname=parsed.get("hostname") or hostname, data=parsed, reason=None, access_status="Success")


# ============================================================
# RESULT HELPERS
# ============================================================

def build_enrichment_result(
    *,
    success: bool = False,
    ip_address: str | None = None,
    mac_address: str | None = None,
    hostname: str | None = None,
    data: dict[str, Any] | None = None,
    reason: str | None = None,
    access_status: str | None = None,
) -> dict[str, Any]:
    """Build a consistent enrichment result."""
    details = data or {}
    cpu_details = details.get("cpu")
    cpu_value = cpu_details.get("name") if isinstance(cpu_details, dict) else cpu_details
    ram_value = details.get("ram")
    if ram_value is None and details.get("ram_gb") is not None:
        try:
            ram_value = f"{float(details['ram_gb']):g} GB"
        except (TypeError, ValueError):
            ram_value = details.get("ram_gb")

    network = details.get("network") or []
    if isinstance(network, dict):
        network = [network]
    primary_adapter = network[0] if network and isinstance(network[0], dict) else {}

    result: dict[str, Any] = {
        "success": success,
        "ip_address": ip_address,
        "mac_address": normalize_mac_address(mac_address),
        "hostname": normalize_host_name(hostname or details.get("hostname")),
        "data": details,
        "reason": reason,
        "access_status": access_status,
        # Compatibility fields consumed by Network Discovery.
        "device_type": details.get("device_type"),
        "manufacturer": details.get("manufacturer"),
        "model": details.get("model"),
        "serial_number": details.get("serial_number"),
        "operating_system": details.get("operating_system"),
        "windows_version": details.get("windows_version"),
        "windows_build": details.get("build_number"),
        "cpu": cpu_value,
        "ram": ram_value,
        "architecture": details.get("architecture"),
        "domain": details.get("domain"),
        "workgroup": details.get("workgroup"),
        "logged_in_user": details.get("logged_in_user"),
        "bios": details.get("bios"),
        "adapter_name": primary_adapter.get("name"),
        "adapter_vendor": primary_adapter.get("vendor"),
        "link_speed": primary_adapter.get("speed"),
        "installed_software": details.get("installed_software"),
        "enrichment_status": "Complete" if success else "Unavailable",
    }

    return result


def _enrichment_reason_from_output(
    stdout: str,
    stderr: str,
) -> str:
    """Generate a readable reason from PowerShell output."""
    combined = f"{stdout}\n{stderr}".strip()

    if not combined:
        return "No response received from remote device."

    lower = combined.lower()

    if (
        "access is denied" in lower
        or "access denied" in lower
        or "unauthorized" in lower
    ):
        return "Access denied while querying the remote device."

    if (
        "rpc server is unavailable" in lower
        or "rpc server unavailable" in lower
    ):
        return "RPC server is unavailable."

    if (
        "winrm" in lower
        and "not configured" in lower
    ):
        return "Remote management is not configured."

    if "timeout" in lower:
        return "Remote query timed out."

    return combined[-1000:]


def _has_known_enrichment_values(
    data: dict[str, Any],
) -> bool:
    """Check whether useful device data was returned."""
    keys = (
        "hostname",
        "manufacturer",
        "model",
        "serial_number",
        "operating_system",
        "windows_version",
        "cpu",
        "ram",
        "bios",
        "network",
        "disks",
        "installed_software",
    )

    for key in keys:
        value = data.get(key)

        if value is None:
            continue

        if isinstance(value, (list, dict)):
            if value:
                return True

        elif str(value).strip():
            return True

    return False


def _summarize_enrichment_status(
    data: dict[str, Any],
) -> str:
    """Create a simple human-readable enrichment status."""
    if not data:
        return "No information collected."

    available: list[str] = []

    if data.get("hostname"):
        available.append("hostname")

    if data.get("manufacturer") or data.get("model"):
        available.append("hardware")

    if data.get("serial_number"):
        available.append("serial")

    if data.get("operating_system") or data.get("windows_version"):
        available.append("OS")

    if data.get("cpu"):
        available.append("CPU")

    if data.get("ram"):
        available.append("RAM")

    if data.get("bios"):
        available.append("BIOS")

    if data.get("network"):
        available.append("network")

    if data.get("disks"):
        available.append("disks")

    if data.get("installed_software"):
        available.append("software")

    if not available:
        return "No known enrichment values."

    return "Collected: " + ", ".join(available)


# ============================================================
# TCP PROBE
# ============================================================

def probe_remote_tcp_port(
    ip_address: str,
    port: int = 135,
    timeout: float = 2.0,
) -> bool:
    """
    Check whether a TCP port is reachable.

    Port 135 is used as the initial check for Windows DCOM/RPC.
    """
    try:
        with socket.create_connection(
            (ip_address, port),
            timeout=timeout,
        ):
            return True

    except Exception:
        return False


# ============================================================
# REMOTE WINDOWS ENRICHMENT
# ============================================================

def query_remote_windows_enrichment(
    ip_address: str,
    mac_address: str | None = None,
    hostname: str | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    """
    Collect detailed information from a remote Windows computer.

    Uses ONE DCOM CIM session.

    Collected information:
        - ComputerSystem
        - OperatingSystem
        - Processor
        - BIOS
        - Network adapters
        - Logical disks
        - Domain/workgroup
        - Logged-in user
        - Windows version/build
        - RAM
        - CPU
        - BIOS information
        - Network adapter vendor/speed
        - Installed software from Windows Registry

    Installed software is collected with StdRegProv instead of
    Win32_Product.
    """

    safe_ip = safe_text(ip_address)
    safe_mac = safe_text(mac_address)
    safe_hostname = safe_text(hostname)

    if not safe_ip:
        return build_enrichment_result(
            success=False,
            ip_address=ip_address,
            mac_address=mac_address,
            hostname=hostname,
            reason="No IP address supplied.",
        )

    # --------------------------------------------------------
    # IMPORTANT:
    # This is intentionally NOT an f-string.
    #
    # PowerShell contains many { } characters. Using an f-string
    # here would cause Pylance errors such as:
    #
    # "Single close brace is not allowed within f-string literal"
    #
    # We therefore use placeholders and replace() below.
    # --------------------------------------------------------

    script = r"""
$ErrorActionPreference = "Stop"

$targetIp = "__TARGET_IP__"
$targetMac = "__TARGET_MAC__"
$targetHostname = "__TARGET_HOSTNAME__"

$session = $null

function Safe-String {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    try {
        return ([string]$Value).Trim()
    }
    catch {
        return ""
    }
}

function Convert-BytesToGB {
    param(
        [object]$Bytes
    )

    if ($null -eq $Bytes) {
        return 0
    }

    try {
        return [math]::Round(
            ([double]$Bytes / 1GB),
            2
        )
    }
    catch {
        return 0
    }
}

function Get-RegistryValue {
    param(
        [object]$CimSession,
        [uint32]$Hive,
        [string]$Path,
        [string]$Name
    )

    try {
        $result = Invoke-CimMethod `
            -CimSession $CimSession `
            -ClassName StdRegProv `
            -MethodName GetStringValue `
            -Arguments @{
                hDefKey = $Hive
                sSubKeyName = $Path
                sValueName = $Name
            } `
            -ErrorAction Stop

        if ($null -ne $result) {
            return Safe-String $result.sValue
        }
    }
    catch {
        return ""
    }

    return ""
}

function Get-InstalledSoftware {
    param(
        [object]$CimSession
    )

    $software = New-Object System.Collections.Generic.List[object]

    $hklm = [uint32]2147483650

    $registryPaths = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($rootPath in $registryPaths) {

        try {
            $keyResult = Invoke-CimMethod `
                -CimSession $CimSession `
                -ClassName StdRegProv `
                -MethodName EnumKey `
                -Arguments @{
                    hDefKey = $hklm
                    sSubKeyName = $rootPath
                } `
                -ErrorAction Stop

            if ($null -eq $keyResult) {
                continue
            }

            $subKeys = $keyResult.sNames

            if ($null -eq $subKeys) {
                continue
            }

            foreach ($subKey in $subKeys) {

                if ([string]::IsNullOrWhiteSpace($subKey)) {
                    continue
                }

                $fullPath = "$rootPath\$subKey"

                $displayName = Get-RegistryValue `
                    -CimSession $CimSession `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "DisplayName"

                if ([string]::IsNullOrWhiteSpace($displayName)) {
                    continue
                }

                $displayVersion = Get-RegistryValue `
                    -CimSession $CimSession `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "DisplayVersion"

                $publisher = Get-RegistryValue `
                    -CimSession $CimSession `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "Publisher"

                $installDate = Get-RegistryValue `
                    -CimSession $CimSession `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "InstallDate"

                $software.Add(
                    [PSCustomObject]@{
                        application = $displayName
                        version = $displayVersion
                        publisher = $publisher
                        install_date = $installDate
                    }
                )
            }
        }
        catch {
            continue
        }
    }

    return @(
        $software |
        Sort-Object application, version, publisher -Unique
    )
}

try {

    # ========================================================
    # CREATE ONE DCOM SESSION
    # ========================================================

    $sessionOption = New-CimSessionOption `
        -Protocol Dcom

    $session = New-CimSession `
        -ComputerName $targetIp `
        -SessionOption $sessionOption `
        -ErrorAction Stop

    # ========================================================
    # COMPUTER SYSTEM
    # ========================================================

    $computer = Get-CimInstance `
        -CimSession $session `
        -ClassName Win32_ComputerSystem `
        -ErrorAction Stop

    # ========================================================
    # OPERATING SYSTEM
    # ========================================================

    $os = Get-CimInstance `
        -CimSession $session `
        -ClassName Win32_OperatingSystem `
        -ErrorAction Stop

    # ========================================================
    # PROCESSOR
    # ========================================================

    $processors = @(
        Get-CimInstance `
            -CimSession $session `
            -ClassName Win32_Processor `
            -ErrorAction Stop
    )

    $primaryProcessor = $processors | Select-Object -First 1

    # ========================================================
    # BIOS
    # ========================================================

    $bios = Get-CimInstance `
        -CimSession $session `
        -ClassName Win32_BIOS `
        -ErrorAction Stop

    # ========================================================
    # NETWORK ADAPTER CONFIGURATION
    # ========================================================

    $networkConfigs = @(
        Get-CimInstance `
            -CimSession $session `
            -ClassName Win32_NetworkAdapterConfiguration `
            -Filter "IPEnabled = True" `
            -ErrorAction Stop
    )

    # ========================================================
    # NETWORK ADAPTERS
    # ========================================================

    $networkAdapters = @(
        Get-CimInstance `
            -CimSession $session `
            -ClassName Win32_NetworkAdapter `
            -ErrorAction Stop
    )

    # ========================================================
    # LOGICAL DISKS
    # ========================================================

    $disks = @(
        Get-CimInstance `
            -CimSession $session `
            -ClassName Win32_LogicalDisk `
            -Filter "DriveType = 3" `
            -ErrorAction Stop
    )

    # ========================================================
    # BASIC COMPUTER INFORMATION
    # ========================================================

    $hostnameResult = Safe-String $computer.Name

    $manufacturer = Safe-String $computer.Manufacturer

    $model = Safe-String $computer.Model

    $serialNumber = Safe-String $bios.SerialNumber

    $domain = Safe-String $computer.Domain

    $workgroup = ""

    if ($computer.PartOfDomain -eq $false) {
        $workgroup = $domain
        $domain = ""
    }

    $loggedInUser = Safe-String $computer.UserName

    $deviceType = "Desktop"

    if (
        $computer.PCSystemType -eq 2
        -or $computer.PCSystemType -eq 8
    ) {
        $deviceType = "Laptop"
    }

    # ========================================================
    # OPERATING SYSTEM INFORMATION
    # ========================================================

    $operatingSystem = Safe-String $os.Caption

    $windowsVersion = Safe-String $os.Version

    $buildNumber = Safe-String $os.BuildNumber

    $architecture = Safe-String $os.OSArchitecture

    $registeredUser = Safe-String $os.RegisteredUser

    $installDate = ""

    try {
        if ($os.InstallDate) {
            $installDate = $os.InstallDate.ToString("yyyy-MM-dd HH:mm:ss")
        }
    }
    catch {
        $installDate = ""
    }

    # ========================================================
    # RAM
    # ========================================================

    $ramGB = Convert-BytesToGB $computer.TotalPhysicalMemory

    # ========================================================
    # CPU
    # ========================================================

    $cpuName = Safe-String $primaryProcessor.Name

    $cpuManufacturer = Safe-String $primaryProcessor.Manufacturer

    $cpuCores = 0

    $cpuThreads = 0

    try {
        $cpuCores = [int]$primaryProcessor.NumberOfCores
    }
    catch {
        $cpuCores = 0
    }

    try {
        $cpuThreads = [int]$primaryProcessor.NumberOfLogicalProcessors
    }
    catch {
        $cpuThreads = 0
    }

    $cpuMaxClockMHz = 0

    try {
        $cpuMaxClockMHz = [int]$primaryProcessor.MaxClockSpeed
    }
    catch {
        $cpuMaxClockMHz = 0
    }

    # ========================================================
    # BIOS
    # ========================================================

    $biosManufacturer = Safe-String $bios.Manufacturer

    $biosVersion = Safe-String $bios.SMBIOSBIOSVersion

    $biosReleaseDate = ""

    try {
        if ($bios.ReleaseDate) {
            $biosReleaseDate = $bios.ReleaseDate.ToString("yyyy-MM-dd")
        }
    }
    catch {
        $biosReleaseDate = ""
    }

    # ========================================================
    # NETWORK INFORMATION
    # ========================================================

    $networkList = New-Object System.Collections.Generic.List[object]

    foreach ($config in $networkConfigs) {

        $adapterName = Safe-String $config.Description

        $mac = Safe-String $config.MACAddress

        $ipAddresses = @()

        if ($null -ne $config.IPAddress) {
            $ipAddresses = @(
                $config.IPAddress |
                Where-Object {
                    $_ -and
                    $_ -notmatch ":"
                }
            )
        }

        $adapterInfo = $networkAdapters |
            Where-Object {
                $_.MACAddress -and
                $mac -and
                $_.MACAddress -eq $mac
            } |
            Select-Object -First 1

        $adapterVendor = ""

        $adapterSpeed = 0

        if ($null -ne $adapterInfo) {

            $adapterVendor = Safe-String $adapterInfo.Manufacturer

            try {
                $adapterSpeed = [double]$adapterInfo.Speed
            }
            catch {
                $adapterSpeed = 0
            }
        }

        $networkList.Add(
            [PSCustomObject]@{
                name = $adapterName
                mac_address = $mac
                ip_addresses = @($ipAddresses)
                gateway = @($config.DefaultIPGateway)
                dns_servers = @($config.DNSServerSearchOrder)
                dhcp_enabled = $config.DHCPEnabled
                vendor = $adapterVendor
                speed = $adapterSpeed
            }
        )
    }

    # ========================================================
    # DISKS
    # ========================================================

    $diskList = New-Object System.Collections.Generic.List[object]

    foreach ($disk in $disks) {

        $sizeGB = Convert-BytesToGB $disk.Size

        $freeGB = Convert-BytesToGB $disk.FreeSpace

        $usedGB = 0

        try {
            $usedGB = [math]::Round(
                ([double]$sizeGB - [double]$freeGB),
                2
            )
        }
        catch {
            $usedGB = 0
        }

        $diskList.Add(
            [PSCustomObject]@{
                drive = Safe-String $disk.DeviceID
                volume_name = Safe-String $disk.VolumeName
                filesystem = Safe-String $disk.FileSystem
                size_gb = $sizeGB
                free_gb = $freeGB
                used_gb = $usedGB
            }
        )
    }

    # ========================================================
    # REMOTE IP ADDRESSES
    # ========================================================

    $remoteIpAddresses = New-Object System.Collections.Generic.List[string]

    foreach ($config in $networkConfigs) {

        if ($null -eq $config.IPAddress) {
            continue
        }

        foreach ($ip in $config.IPAddress) {

            if (
                $ip -and
                $ip -notmatch ":"
            ) {
                if (
                    $ip -notin $remoteIpAddresses
                ) {
                    $remoteIpAddresses.Add($ip)
                }
            }
        }
    }

    # ========================================================
    # INSTALLED SOFTWARE
    #
    # IMPORTANT:
    # We use the SAME DCOM CIM session here.
    #
    # No second New-CimSession.
    # No Win32_Product.
    # ========================================================

    $installedSoftware = @(
        Get-InstalledSoftware `
            -CimSession $session
    )

    # ========================================================
    # SOFTWARE ACCESS STATUS
    # ========================================================

    $softwareAccessStatus = "Available"

    if (
        $null -eq $installedSoftware
    ) {
        $installedSoftware = @()

        $softwareAccessStatus = "Unavailable"
    }

    # ========================================================
    # BUILD FINAL RESULT
    # ========================================================

    $result = [PSCustomObject]@{

        success = $true

        access_status = "Success"

        software_access_status = $softwareAccessStatus

        ip_address = $targetIp

        mac_address = $targetMac

        hostname = $hostnameResult

        manufacturer = $manufacturer

        model = $model

        serial_number = $serialNumber

        device_type = $deviceType

        domain = $domain

        workgroup = $workgroup

        logged_in_user = $loggedInUser

        operating_system = $operatingSystem

        windows_version = $windowsVersion

        build_number = $buildNumber

        architecture = $architecture

        registered_user = $registeredUser

        os_install_date = $installDate

        ram_gb = $ramGB

        cpu = [PSCustomObject]@{
            name = $cpuName
            manufacturer = $cpuManufacturer
            cores = $cpuCores
            logical_processors = $cpuThreads
            max_clock_mhz = $cpuMaxClockMHz
        }

        bios = [PSCustomObject]@{
            manufacturer = $biosManufacturer
            version = $biosVersion
            serial_number = $serialNumber
            release_date = $biosReleaseDate
        }

        network = @($networkList)

        ip_addresses = @($remoteIpAddresses)

        disks = @($diskList)

        installed_software = @($installedSoftware)

        installed_software_count = @(
            $installedSoftware
        ).Count

        enrichment_status = "Complete"
    }

    $result | ConvertTo-Json -Depth 10 -Compress
}
catch {

    $errorMessage = Safe-String $_.Exception.Message

    $accessStatus = "Error"

    if (
        $errorMessage -match "Access is denied"
        -or
        $errorMessage -match "Access denied"
        -or
        $errorMessage -match "Unauthorized"
    ) {
        $accessStatus = "Access Denied"
    }

    elseif (
        $errorMessage -match "RPC server is unavailable"
        -or
        $errorMessage -match "RPC"
    ) {
        $accessStatus = "RPC Unavailable"
    }

    $errorResult = [PSCustomObject]@{

        success = $false

        access_status = $accessStatus

        software_access_status = "Unavailable"

        ip_address = $targetIp

        mac_address = $targetMac

        hostname = $targetHostname

        reason = $errorMessage

        installed_software = @()

        installed_software_count = 0
    }

    $errorResult | ConvertTo-Json -Depth 10 -Compress
}
finally {

    if ($null -ne $session) {

        try {
            Remove-CimSession `
                -CimSession $session `
                -ErrorAction SilentlyContinue
        }
        catch {
        }
    }
}
"""

    # --------------------------------------------------------
    # Inject values safely.
    #
    # We do NOT use an f-string.
    # --------------------------------------------------------

    safe_ip_ps = safe_ip.replace(
        "\\",
        "\\\\",
    ).replace(
        "'",
        "''",
    )

    safe_mac_ps = safe_mac.replace(
        "\\",
        "\\\\",
    ).replace(
        "'",
        "''",
    )

    safe_hostname_ps = safe_hostname.replace(
        "\\",
        "\\\\",
    ).replace(
        "'",
        "''",
    )

    script = (
        script
        .replace(
            "__TARGET_IP__",
            safe_ip_ps,
        )
        .replace(
            "__TARGET_MAC__",
            safe_mac_ps,
        )
        .replace(
            "__TARGET_HOSTNAME__",
            safe_hostname_ps,
        )
    )

    # --------------------------------------------------------
    # Run PowerShell
    # --------------------------------------------------------

    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]

    return_code, stdout, stderr = run_windows_command(
        command,
        timeout=timeout,
    )

    if return_code != 0:
        reason = _enrichment_reason_from_output(
            stdout,
            stderr,
        )

        return build_enrichment_result(
            success=False,
            ip_address=safe_ip,
            mac_address=safe_mac,
            hostname=safe_hostname,
            reason=reason,
            access_status=(
                "Access Denied"
                if "access denied" in reason.lower()
                else "Error"
            ),
        )

    parsed = parse_powershell_json(stdout)

    if not isinstance(parsed, dict):
        return build_enrichment_result(
            success=False,
            ip_address=safe_ip,
            mac_address=safe_mac,
            hostname=safe_hostname,
            reason=(
                "PowerShell returned an invalid JSON response."
            ),
            access_status="Invalid Response",
        )

    # --------------------------------------------------------
    # Normalize installed software
    # --------------------------------------------------------

    installed_software = parsed.get(
        "installed_software",
        [],
    )

    if installed_software is None:
        installed_software = []

    if isinstance(installed_software, dict):
        installed_software = [
            installed_software
        ]

    if not isinstance(installed_software, list):
        installed_software = []

    parsed["installed_software"] = installed_software

    parsed["installed_software_count"] = len(
        installed_software
    )

    # --------------------------------------------------------
    # Make sure input values are preserved if the remote
    # machine did not return them.
    # --------------------------------------------------------

    if not parsed.get("ip_address"):
        parsed["ip_address"] = safe_ip

    if not parsed.get("mac_address"):
        parsed["mac_address"] = safe_mac

    if not parsed.get("hostname"):
        parsed["hostname"] = safe_hostname

    success = bool(
        parsed.get("success", False)
    )

    access_status = safe_text(
        parsed.get(
            "access_status",
            "",
        )
    )

    reason = parsed.get("reason")

    if not success and not reason:
        reason = (
            "Remote Windows enrichment failed."
        )

    return build_enrichment_result(
        success=success,
        ip_address=parsed.get(
            "ip_address",
            safe_ip,
        ),
        mac_address=parsed.get(
            "mac_address",
            safe_mac,
        ),
        hostname=parsed.get(
            "hostname",
            safe_hostname,
        ),
        data=parsed,
        reason=reason,
        access_status=access_status,
    )


# ============================================================
# REMOTE INSTALLED SOFTWARE
# ============================================================

def get_remote_installed_software(
    ip_address: str,
    timeout: int = 60,
) -> list[dict[str, Any]]:
    """
    Compatibility function for remote installed software.

    The normal enrichment flow does NOT need to call this
    separately because query_remote_windows_enrichment()
    already collects installed software using the same DCOM
    session.

    This function is kept in case another part of the project
    still calls it.
    """

    safe_ip = safe_text(ip_address)

    if not safe_ip:
        return []

    script = r"""
$ErrorActionPreference = "Stop"

$targetIp = "__TARGET_IP__"

$session = $null

function Safe-String {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    try {
        return ([string]$Value).Trim()
    }
    catch {
        return ""
    }
}

function Get-RegistryValue {
    param(
        [object]$CimSession,
        [uint32]$Hive,
        [string]$Path,
        [string]$Name
    )

    try {

        $result = Invoke-CimMethod `
            -CimSession $CimSession `
            -ClassName StdRegProv `
            -MethodName GetStringValue `
            -Arguments @{
                hDefKey = $Hive
                sSubKeyName = $Path
                sValueName = $Name
            } `
            -ErrorAction Stop

        if ($null -ne $result) {
            return Safe-String $result.sValue
        }
    }
    catch {
        return ""
    }

    return ""
}

try {

    # ========================================================
    # ONE DCOM SESSION
    # ========================================================

    $sessionOption = New-CimSessionOption `
        -Protocol Dcom

    $session = New-CimSession `
        -ComputerName $targetIp `
        -SessionOption $sessionOption `
        -ErrorAction Stop

    $hklm = [uint32]2147483650

    $registryPaths = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $software = New-Object System.Collections.Generic.List[object]

    foreach ($rootPath in $registryPaths) {

        try {

            $keyResult = Invoke-CimMethod `
                -CimSession $session `
                -ClassName StdRegProv `
                -MethodName EnumKey `
                -Arguments @{
                    hDefKey = $hklm
                    sSubKeyName = $rootPath
                } `
                -ErrorAction Stop

            if ($null -eq $keyResult) {
                continue
            }

            if ($null -eq $keyResult.sNames) {
                continue
            }

            foreach ($subKey in $keyResult.sNames) {

                if ([string]::IsNullOrWhiteSpace($subKey)) {
                    continue
                }

                $fullPath = "$rootPath\$subKey"

                $displayName = Get-RegistryValue `
                    -CimSession $session `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "DisplayName"

                if ([string]::IsNullOrWhiteSpace($displayName)) {
                    continue
                }

                $displayVersion = Get-RegistryValue `
                    -CimSession $session `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "DisplayVersion"

                $publisher = Get-RegistryValue `
                    -CimSession $session `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "Publisher"

                $installDate = Get-RegistryValue `
                    -CimSession $session `
                    -Hive $hklm `
                    -Path $fullPath `
                    -Name "InstallDate"

                $software.Add(
                    [PSCustomObject]@{
                        application = $displayName
                        version = $displayVersion
                        publisher = $publisher
                        install_date = $installDate
                    }
                )
            }
        }
        catch {
            continue
        }
    }

    @(
        $software |
        Sort-Object application, version, publisher -Unique
    ) |
    ConvertTo-Json -Depth 10 -Compress
}
catch {

    @() |
    ConvertTo-Json -Depth 10 -Compress
}
finally {

    if ($null -ne $session) {

        try {
            Remove-CimSession `
                -CimSession $session `
                -ErrorAction SilentlyContinue
        }
        catch {
        }
    }
}
"""

    script = script.replace(
        "__TARGET_IP__",
        safe_ip,
    )

    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]

    return_code, stdout, stderr = run_windows_command(
        command,
        timeout=timeout,
    )

    if return_code != 0:
        logger.debug(
            "Remote software query failed for %s: %s",
            safe_ip,
            stderr,
        )

        return []

    parsed = parse_powershell_json(stdout)

    if parsed is None:
        return []

    if isinstance(parsed, dict):
        return [parsed]

    if isinstance(parsed, list):
        return [
            item
            for item in parsed
            if isinstance(item, dict)
        ]

    return []


# ============================================================
# POWERSHELL JSON
# ============================================================

def parse_powershell_json(
    output: str,
) -> Any:
    """
    Parse PowerShell JSON output.

    PowerShell can sometimes produce extra whitespace or
    diagnostic output, so we try the complete output first and
    then search for the last JSON object/array.
    """
    if not output:
        return None

    text = output.strip()

    if not text:
        return None

    # --------------------------------------------------------
    # First attempt: complete output
    # --------------------------------------------------------

    try:
        return json.loads(text)

    except json.JSONDecodeError:
        pass

    # --------------------------------------------------------
    # Try line by line from the bottom.
    # --------------------------------------------------------

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    for line in reversed(lines):

        if not (
            line.startswith("{")
            or line.startswith("[")
        ):
            continue

        try:
            return json.loads(line)

        except json.JSONDecodeError:
            continue

    # --------------------------------------------------------
    # Last attempt: find JSON object/array boundaries.
    # --------------------------------------------------------

    object_start = text.find("{")
    object_end = text.rfind("}")

    if (
        object_start >= 0
        and object_end > object_start
    ):
        candidate = text[
            object_start:object_end + 1
        ]

        try:
            return json.loads(candidate)

        except json.JSONDecodeError:
            pass

    array_start = text.find("[")
    array_end = text.rfind("]")

    if (
        array_start >= 0
        and array_end > array_start
    ):
        candidate = text[
            array_start:array_end + 1
        ]

        try:
            return json.loads(candidate)

        except json.JSONDecodeError:
            pass

    return None


# ============================================================
# LOCAL INSTALLED SOFTWARE
# ============================================================

def get_local_installed_software() -> list[dict[str, Any]]:
    """
    Collect installed software from the local Windows machine.

    Uses Windows Registry instead of Win32_Product.
    """

    script = r"""
$ErrorActionPreference = "SilentlyContinue"

function Get-RegistryValue {
    param(
        [string]$Path,
        [string]$Name
    )

    try {
        $value = Get-ItemPropertyValue `
            -Path $Path `
            -Name $Name `
            -ErrorAction Stop

        if ($null -ne $value) {
            return ([string]$value).Trim()
        }
    }
    catch {
        return ""
    }

    return ""
}

$paths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$software = New-Object System.Collections.Generic.List[object]

foreach ($path in $paths) {

    try {

        $items = Get-ItemProperty `
            -Path $path `
            -ErrorAction SilentlyContinue

        foreach ($item in $items) {

            $displayName = ""

            try {
                $displayName = ([string]$item.DisplayName).Trim()
            }
            catch {
                $displayName = ""
            }

            if ([string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            $version = ""

            try {
                $version = ([string]$item.DisplayVersion).Trim()
            }
            catch {
                $version = ""
            }

            $publisher = ""

            try {
                $publisher = ([string]$item.Publisher).Trim()
            }
            catch {
                $publisher = ""
            }

            $installDate = ""

            try {
                $installDate = ([string]$item.InstallDate).Trim()
            }
            catch {
                $installDate = ""
            }

            $software.Add(
                [PSCustomObject]@{
                    application = $displayName
                    version = $version
                    publisher = $publisher
                    install_date = $installDate
                }
            )
        }
    }
    catch {
        continue
    }
}

@(
    $software |
    Sort-Object application, version, publisher -Unique
) |
ConvertTo-Json -Depth 10 -Compress
"""

    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]

    return_code, stdout, stderr = run_windows_command(
        command,
        timeout=60,
    )

    if return_code != 0:
        logger.debug(
            "Local software query failed: %s",
            stderr,
        )

        return []

    parsed = parse_powershell_json(stdout)

    if parsed is None:
        return []

    if isinstance(parsed, dict):
        return [parsed]

    if isinstance(parsed, list):
        return [
            item
            for item in parsed
            if isinstance(item, dict)
        ]

    return []


# ============================================================
# MAIN ENRICHMENT FUNCTION
# ============================================================

def collect_device_enrichment(
    ip_address: str,
    mac_address: str | None = None,
    hostname: str | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    """
    Main device enrichment entry point.

    The Network Discovery system provides the IP address.

    Flow:

        Network Discovery
                ↓
        collect_device_enrichment()
                ↓
        Check whether target is local
                ↓
        TCP 135 probe
                ↓
        query_remote_windows_enrichment()
                ↓
        ONE DCOM CIM session
                ↓
        WMI + Registry software
                ↓
        Return JSON/data
    """

    safe_ip = safe_text(ip_address)

    if not safe_ip:
        return build_enrichment_result(
            success=False,
            ip_address=ip_address,
            mac_address=mac_address,
            hostname=hostname,
            reason="No IP address supplied.",
        )

    # --------------------------------------------------------
    # Determine whether this is the local computer.
    # --------------------------------------------------------

    local_ips = extract_local_ip_addresses()

    is_local = safe_ip in local_ips

    # --------------------------------------------------------
    # LOCAL DEVICE
    # --------------------------------------------------------

    if is_local:
        return query_local_windows_enrichment(
            safe_ip,
            mac_address=mac_address,
            hostname=hostname,
            timeout=timeout,
        )

    # --------------------------------------------------------
    # REMOTE DEVICE
    # --------------------------------------------------------

    # First check TCP 135.
    #
    # Network Discovery has already found the IP, but this
    # avoids launching PowerShell when RPC/DCOM is obviously
    # unreachable.

    if not probe_remote_tcp_port(
        safe_ip,
        port=135,
        timeout=2.0,
    ):
        return build_enrichment_result(
            success=False,
            ip_address=safe_ip,
            mac_address=mac_address,
            hostname=hostname,
            reason=(
                "TCP port 135 is not reachable. "
                "Remote DCOM/RPC enrichment cannot start."
            ),
            access_status="RPC Unavailable",
        )

    # --------------------------------------------------------
    # Remote enrichment.
    #
    # IMPORTANT:
    #
    # Do NOT call get_remote_installed_software() here.
    #
    # query_remote_windows_enrichment() already collects:
    #
    #   WMI information
    #   +
    #   Registry installed software
    #
    # through ONE DCOM session.
    # --------------------------------------------------------

    result = query_remote_windows_enrichment(
        safe_ip,
        mac_address=mac_address,
        hostname=hostname,
        timeout=timeout,
    )

    return result

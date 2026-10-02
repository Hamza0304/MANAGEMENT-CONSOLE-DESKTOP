# ============================================
# Windows PC Inventory Collector
# ============================================

$OutputFolder = "C:\Users\Hamza Senhaji\OneDrive\Desktop\test 1"

# Create output folder if it does not exist
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$PCName = $env:COMPUTERNAME
$UserName = $env:USERNAME

Write-Host "Collecting inventory for $PCName ..."

# --------------------------------------------
# Operating System
# --------------------------------------------

$OS = Get-CimInstance Win32_OperatingSystem

# --------------------------------------------
# Computer information
# --------------------------------------------

$Computer = Get-CimInstance Win32_ComputerSystem

# --------------------------------------------
# CPU
# --------------------------------------------

$CPU = Get-CimInstance Win32_Processor |
    Select-Object -First 1

# --------------------------------------------
# BIOS / Serial Number
# --------------------------------------------

$BIOS = Get-CimInstance Win32_BIOS

# --------------------------------------------
# IP address
# --------------------------------------------

$IPAddresses = Get-NetIPAddress -AddressFamily IPv4 |
    Where-Object {
        $_.IPAddress -notlike "169.*" -and
        $_.IPAddress -ne "127.0.0.1"
    } |
    Select-Object -ExpandProperty IPAddress

$IP = $IPAddresses -join ", "

# --------------------------------------------
# MAC address
# --------------------------------------------

$MACAddresses = Get-CimInstance Win32_NetworkAdapterConfiguration |
    Where-Object {
        $_.IPEnabled -eq $true
    } |
    Select-Object -ExpandProperty MACAddress

$MAC = $MACAddresses -join ", "

# --------------------------------------------
# RAM
# --------------------------------------------

$RAMGB = [math]::Round(
    $Computer.TotalPhysicalMemory / 1GB,
    2
)

# --------------------------------------------
# BitLocker
# --------------------------------------------

$BitLockerStatus = "Unknown"

try {

    $BitLockerVolumes = Get-BitLockerVolume -ErrorAction Stop

    if ($BitLockerVolumes) {

        $ProtectedVolumes = $BitLockerVolumes |
            Where-Object {
                $_.ProtectionStatus -eq 1
            }

        if ($ProtectedVolumes) {
            $BitLockerStatus = "Chiffré"
        }
        else {
            $BitLockerStatus = "Non chiffré"
        }
    }

}
catch {

    $BitLockerStatus = "Unavailable"
}

# --------------------------------------------
# PC inventory object
# --------------------------------------------

$PCInfo = [PSCustomObject]@{

    PCName       = $PCName
    User         = $UserName
    OS           = $OS.Caption
    OSVersion    = $OS.Version
    RAM_GB       = $RAMGB
    CPU          = $CPU.Name
    IP           = $IP
    MAC          = $MAC
    SerialNumber = $BIOS.SerialNumber
    Model        = $Computer.Model
    BitLocker    = $BitLockerStatus
}

# --------------------------------------------
# Export PC inventory
# --------------------------------------------

$InventoryFile = Join-Path $OutputFolder "inventaire_parc.csv"

$PCInfo |
    Export-Csv `
        -Path $InventoryFile `
        -NoTypeInformation `
        -Encoding UTF8

Write-Host "PC inventory saved to:"
Write-Host $InventoryFile

# --------------------------------------------
# Installed applications
# --------------------------------------------

$Apps = @()

$RegistryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($Path in $RegistryPaths) {

    try {

        $Entries = Get-ItemProperty $Path -ErrorAction SilentlyContinue

        foreach ($Entry in $Entries) {

            if ($Entry.DisplayName) {

                $Apps += [PSCustomObject]@{

                    Name            = $Entry.DisplayName
                    Version         = $Entry.DisplayVersion
                    Publisher       = $Entry.Publisher
                    InstallLocation = $Entry.InstallLocation
                }
            }
        }
    }
    catch {
        # Ignore registry errors
    }
}

# Remove duplicate applications
$Apps = $Apps |
    Sort-Object Name, Version -Unique

# --------------------------------------------
# Export applications
# --------------------------------------------

$AppsFile = Join-Path $OutputFolder "Apps_$PCName.csv"

$Apps |
    Export-Csv `
        -Path $AppsFile `
        -NoTypeInformation `
        -Encoding UTF8

Write-Host "Installed applications saved to:"
Write-Host $AppsFile

Write-Host ""
Write-Host "Inventory collection completed successfully."
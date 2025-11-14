Set-ExecutionPolicy Bypass -Scope Process -Force

$TargetDeviceName = "USB"

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Write-Host "Searching for USB device matching: $TargetDeviceName ..." -ForegroundColor Cyan
$device = Get-PnpDevice | Where-Object { $_.FriendlyName -like "*$TargetDeviceName*" -and $_.Status -eq "OK" } | Select-Object -First 1

if (-not $device) {
    Write-Error "No matching USB device found."
    exit 1
}

Write-Host "Found device: $($device.FriendlyName) [$($device.InstanceId)]" -ForegroundColor Green

$driverInfo = pnputil /enum-drivers | Select-String -Context 0,4 $device.InstanceId
if (-not $driverInfo) {
    Write-Warning "Could not find driver package for this device. Proceeding with uninstall."
}

Write-Host "Uninstalling device..." -ForegroundColor Yellow
try {
    Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop
    Uninstall-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop
    Write-Host "Device uninstalled successfully." -ForegroundColor Green
} catch {
    Write-Error "Failed to uninstall device: $_"
    exit 1
}

if ($driverInfo) {
    $oemLine = ($driverInfo.Context.PostContext | Select-String "Published Name").ToString()
    if ($oemLine -match "oem\d+\.inf") {
        $oemFile = $matches[0]
        Write-Host "Removing driver package: $oemFile ..." -ForegroundColor Yellow
        pnputil /delete-driver $oemFile /uninstall /force
    }
}

Write-Host "Triggering Windows Update driver search..." -ForegroundColor Cyan
try {
    pnputil /scan-devices
    Start-Process "ms-settings:windowsupdate"
    Write-Host "Windows will now attempt to reinstall the driver from Windows Update." -ForegroundColor Green
    Write-Host "If it doesn't install automatically, click 'Check for updates' in the Windows Update window."
} catch {
    Write-Error "Failed to trigger driver reinstall: $_"
}

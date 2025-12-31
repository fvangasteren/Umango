# Define the registry path
$regPath = "HKLM:\SOFTWARE\Umango"

# Ensure the key exists
if (-not (Test-Path $regPath)) {
    New-Item -Path "HKLM:\SOFTWARE" -Name "Umango" -Force | Out-Null
}

# Create or update the string value
New-ItemProperty -Path $regPath -Name "logging.file.enabled" -Value "true" -PropertyType String -Force | Out-Null

Write-Output "Registry key 'logging.file.enabled' set to 'true' under $regPath"

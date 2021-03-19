# Move up 1 directory level
Set-Location (Get-Location | Split-Path)

# Check if folder exist, delete if yes
if (Test-Path -Path "__pycache__") {
    Remove-Item "__pycache__" -Recurse
}
if (Test-Path -Path "data\app") {
    Remove-Item "data\app" -Recurse
}

# Check if file exist, delete if yes
if (Test-Path "*.log") {
    Remove-Item "*.log"
}

Read-Host "Completed!`nPress any key to continue..."
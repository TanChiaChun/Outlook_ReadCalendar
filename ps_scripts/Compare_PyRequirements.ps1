# Move up 1 directory level
Set-Location (Get-Location | Split-Path)

# Check if requirements file exist, generate Python requirements if no
if (-not (Test-Path "requirements\requirements.txt") ) {
    ps_scripts\Get_PyRequirements.ps1 True
}

# Parse requirements.txt into hash tables
$req_list_source = Get-Content -Path "requirements\requirements.txt"
$req_hash_source = @{}
foreach ($req in $req_list_source) {
    $req_details = $req -split "=="
    $req_hash_source[$req_details[0]] = $req_details[1]
}

# Navigate to default Python packages folder, manual input if not exists
$dest_folder = Join-Path -Path (Get-Location | Split-Path | Split-Path) -ChildPath "Python_Packages"
if (-not (Test-Path -Path $dest_folder) ) {
    $dest_folder = Read-Host "Input source folder for Python packages"
}

# Parse packages folders into hash tables, exclude for older Python versions
$folders_list_dest = Get-ChildItem -Path $dest_folder -Name
$req_hash_dest = @{}
foreach ($folder in $folders_list_dest) {
    $folder_details = $folder -split '_'
    if ($folder_details.Count -eq 2){
        $req_hash_dest[$folder_details[0]] = $folder_details[1]
    }
}

# Compare and output outdated packages
foreach ($key in $req_hash_source.Keys) {
    if ($req_hash_dest.ContainsKey($key)) {
        if ($req_hash_source.$key -lt $req_hash_dest.$key) {
            Write-Host ($key + ": " + $req_hash_source.$key + " -> " + $req_hash_dest.$key)
        }
    }
}

Read-Host "Completed!`nPress any key to continue..."
# Move up 1 directory level
Set-Location (Get-Location | Split-Path)

function My-Clear-Folder {
    # Check if folders exist, delete if yes
    if (Test-Path -Path "__pycache__") {
        Remove-Item "__pycache__" -Recurse
    }
    if (Test-Path -Path "pyinstaller\build") {
        Remove-Item "pyinstaller\build" -Recurse
    }
    if (Test-Path -Path "pyinstaller\dist") {
        Remove-Item "pyinstaller\dist" -Recurse
    }

    # Check if files exist, delete if yes
    if (Test-Path "pyinstaller\*.spec") {
        Remove-Item "pyinstaller\*.spec"
    }
}

# Run My-Clear-Folder function
My-Clear-Folder

# Check if files exist, delete if yes
if (Test-Path "pyinstaller\*.exe") {
    Remove-Item "pyinstaller\*.exe"
}

# Set exe name
$proj_name = Split-Path (Get-Location) -Leaf
$script_name = "$proj_name.py"
if (-not (Test-Path $script_name) ) {
    $proj_name = Read-Host "Input project name"
    $script_name = "$proj_name.py"
}
$proj_version = git describe
if (-not $proj_version) {
    $proj_version = Read-Host "Input project version"
}
$exe_name = "$proj_name" + "_" + "$proj_version.exe"

# Check if folders exist, create if no
if (-not (Test-Path -Path "pyinstaller") ) {
    New-Item -Name "pyinstaller" -ItemType "directory"
}

# Check if files exist, create if no
if (-not (Test-Path "pyinstaller\python-icon.ico") ) {
    $ico_source = Get-Location | Split-Path
    if (-not (Test-Path -Path "$ico_source\python-icon.ico") ) {
        $ico_source = Read-Host "Input source folder for ico file"
    }
    Copy-Item "$ico_source\python-icon.ico" -Destination "pyinstaller"
}

# Run Pyinstaller
venv\Scripts\Activate.ps1
pyinstaller $script_name --onefile --icon python-icon.ico --distpath ./pyinstaller/dist --workpath ./pyinstaller/build --specpath ./pyinstaller
Move-Item -Path "pyinstaller\dist\$proj_name.exe" -Destination "pyinstaller\$exe_name"
My-Clear-Folder

Read-Host "Completed!`nPress any key to continue..."
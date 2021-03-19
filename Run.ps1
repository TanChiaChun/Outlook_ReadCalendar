# Set script name
$proj_name = Split-Path (Get-Location) -Leaf
$script_name = "$proj_name.py"
if (-not (Test-Path $script_name) ) {
    $proj_name = Read-Host "Input project name"
    $script_name = "$proj_name.py"
}

venv\Scripts\python $script_name

Read-Host "Completed!`nPress any key to continue..."
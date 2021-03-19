# Move up 1 directory level
Set-Location (Get-Location | Split-Path)

# Read input
$dest_dir = Read-Host "Input destination folder"

# Generate Python requirements
ps_scripts\Get_PyRequirements.ps1 True

# Check if destination folder exist, delete if yes
if (Test-Path -Path "$dest_dir\.git") {
    Remove-Item "$dest_dir\.git" -Recurse -Force
}
# Check if destination folder exist, create if no
if (-not (Test-Path -Path "$dest_dir\requirements") ) {
    New-Item -Path $dest_dir -Name "requirements" -ItemType "directory"
}
if (-not (Test-Path -Path "$dest_dir\ps_scripts") ) {
    New-Item -Path $dest_dir -Name "ps_scripts" -ItemType "directory"
}

# Copy files
Copy-Item "*.py" -Destination $dest_dir
Copy-Item "requirements\requirements*.txt" -Destination "$dest_dir\requirements"
Copy-Item ".gitignore" -Destination $dest_dir
Copy-Item "*.ps1" -Destination $dest_dir
Copy-Item "ps_scripts\*.ps1" -Destination "$dest_dir\ps_scripts"

# Copy folders
Copy-Item -Path ".git" -Destination $dest_dir -Recurse

# Change folder attribute to hidden
$temp_folder = Get-Item $dest_dir\.git
$temp_folder.Attributes = "Hidden"

Read-Host "Completed!`nPress any key to continue..."
$ErrorActionPreference = "Stop"

$sourceDir = "C:\Users\Nam Phương\OneDrive\Documents\Xóa highlight đề"
$targetDir = "C:\XoaHighlightDe"

Write-Host "1. Creating target directory..."
if (!(Test-Path -Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir | Out-Null
}

Write-Host "2. Copying source code..."
if (!(Test-Path -Path "$targetDir\backend")) {
    New-Item -ItemType Directory -Path "$targetDir\backend" | Out-Null
}
Copy-Item -Path "$sourceDir\backend\*" -Destination "$targetDir\backend" -Recurse -Force -Exclude "venv", "__pycache__"

if (!(Test-Path -Path "$targetDir\frontend")) {
    New-Item -ItemType Directory -Path "$targetDir\frontend" | Out-Null
}
Copy-Item -Path "$sourceDir\frontend\*" -Destination "$targetDir\frontend" -Recurse -Force -Exclude "node_modules"

Write-Host "3. Installing backend dependencies..."
Set-Location -Path "$targetDir\backend"
python -m venv venv
& ".\venv\Scripts\python.exe" -m pip install --upgrade pip
& ".\venv\Scripts\pip.exe" install PyMuPDF==1.24.1
& ".\venv\Scripts\pip.exe" install -r requirements.txt

Write-Host "4. Installing frontend dependencies..."
Set-Location -Path "$targetDir\frontend"
npm install

Write-Host "Done!"

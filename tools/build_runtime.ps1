# Builds the self-contained runtime\ folder (maintainers only - end users never run this).
# Copies a trimmed Python 3.12 + pywin32 from an existing installation; nothing is downloaded and nothing is
# installed on the machine. The project then runs with runtime\python.exe only.
param([string]$Source = "")

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$dest = Join-Path $root "runtime"

if (-not $Source) { $Source = (& python -c "import sys; print(sys.base_prefix)").Trim() }
if (-not (Test-Path "$Source\python.exe")) { throw "No python.exe under $Source" }
$sp = "$Source\Lib\site-packages"
foreach ($need in "win32", "win32com", "pythoncom.py", "pywin32_system32") {
    if (-not (Test-Path "$sp\$need")) { throw "pywin32 part missing in the source Python: $need" }
}

if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
New-Item -ItemType Directory $dest | Out-Null

# interpreter + runtime DLLs (pywin32's DLLs sit next to python.exe, exactly like a normal install)
foreach ($f in "python.exe", "pythonw.exe", "python312.dll", "python3.dll", "vcruntime140.dll", "vcruntime140_1.dll",
               "pythoncom312.dll", "pywintypes312.dll", "LICENSE.txt") {
    Copy-Item "$Source\$f" $dest
}

# extension modules (drop Tk)
robocopy "$Source\DLLs" "$dest\DLLs" /E /XF "_tkinter.pyd" "tcl*.dll" "tk*.dll" /NFL /NDL /NJH /NJS | Out-Null

# standard library without tests / GUI / installers / third-party
robocopy "$Source\Lib" "$dest\Lib" /E /XD "site-packages" "test" "tests" "idlelib" "tkinter" "turtledemo" "ensurepip" `
    "__pycache__" "lib2to3" "pydoc_data" /XF "turtle.py" "*.pyc" /NFL /NDL /NJH /NJS | Out-Null

# only the third-party packages this project imports
$dsp = "$dest\Lib\site-packages"
New-Item -ItemType Directory $dsp | Out-Null
foreach ($d in "win32", "win32com", "win32comext", "pywin32_system32") {
    robocopy "$sp\$d" "$dsp\$d" /E /XD "__pycache__" "test" "tests" /NFL /NDL /NJH /NJS | Out-Null
}
foreach ($f in "pythoncom.py", "pywin32.pth", "pywin32.version.txt", "pywin32_testutil.py") {
    if (Test-Path "$sp\$f") { Copy-Item "$sp\$f" $dsp }
}
Get-ChildItem $sp -Filter "pywin32_bootstrap*" | ForEach-Object { Copy-Item $_.FullName $dsp -Recurse }

# pre-compile so the first start is fast even when the folder is read-only
& "$dest\python.exe" -I -m compileall -q "$dest\Lib" | Out-Null

$mb = [math]::Round((Get-ChildItem $dest -Recurse -File | Measure-Object Length -Sum).Sum / 1MB, 1)
Write-Host "runtime built: $dest  ($mb MB)"

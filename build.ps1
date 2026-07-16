$ErrorActionPreference = "Stop"
$Python = if ($env:BUNDLE_SPLITTER_PYTHON) { $env:BUNDLE_SPLITTER_PYTHON } else { "python" }

if (-not (Test-Path -LiteralPath $Python)) {
    $Python = "python"
}

& $Python -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) { throw "Tests failed with exit code $LASTEXITCODE" }
& $Python -m PyInstaller --noconfirm --clean --onefile --console `
    --name "SplitterWorker" `
    --distpath "work\worker-dist" `
    --workpath "work\worker-build" `
    --specpath "work" `
    "splitter_worker.py"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

$Csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path -LiteralPath $Csc)) { throw "C# compiler not found: $Csc" }
& $Csc /nologo /target:winexe /optimize+ /out:"outputs\BundleUnitPriceSplitter_v3.exe" `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll `
    /resource:"work\worker-dist\SplitterWorker.exe,SplitterWorker.exe" `
    "DesktopApp.cs"
if ($LASTEXITCODE -ne 0) { throw "C# compiler failed with exit code $LASTEXITCODE" }

Write-Host "Built outputs\BundleUnitPriceSplitter_v3.exe"

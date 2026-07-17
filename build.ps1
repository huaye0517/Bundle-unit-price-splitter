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
$LauncherBuildDir = Join-Path $env:TEMP "BundleUnitPriceSplitterBuild"
New-Item -ItemType Directory -Path $LauncherBuildDir -Force | Out-Null
New-Item -ItemType Directory -Path "outputs" -Force | Out-Null
Copy-Item -LiteralPath "DesktopApp.cs" -Destination (Join-Path $LauncherBuildDir "DesktopApp.cs") -Force
Copy-Item -LiteralPath "work\worker-dist\SplitterWorker.exe" -Destination (Join-Path $LauncherBuildDir "SplitterWorker.exe") -Force
$LauncherExe = Join-Path $LauncherBuildDir "BundleUnitPriceSplitter_v8.exe"
& $Csc /nologo /target:winexe /optimize+ /out:$LauncherExe `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll `
    /resource:"$LauncherBuildDir\SplitterWorker.exe,SplitterWorker.exe" `
    "$LauncherBuildDir\DesktopApp.cs"
if ($LASTEXITCODE -ne 0) { throw "C# compiler failed with exit code $LASTEXITCODE" }
Copy-Item -LiteralPath $LauncherExe -Destination "outputs\BundleUnitPriceSplitter_v8.exe" -Force

Write-Host "Built outputs\BundleUnitPriceSplitter_v8.exe"

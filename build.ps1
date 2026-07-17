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
$LauncherExe = Join-Path $LauncherBuildDir "BundleUnitPriceSplitter.exe"
& $Csc /nologo /target:winexe /optimize+ /out:$LauncherExe `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll `
    /resource:"$LauncherBuildDir\SplitterWorker.exe,SplitterWorker.exe" `
    "$LauncherBuildDir\DesktopApp.cs"
if ($LASTEXITCODE -ne 0) { throw "C# compiler failed with exit code $LASTEXITCODE" }
$OutputExeName = (-join @(
    [char]0x7EC4, [char]0x5408, [char]0x88C5, [char]0x5355, [char]0x4EF7,
    [char]0x62C6, [char]0x5206, [char]0x5DE5, [char]0x5177
)) + ".exe"
$OutputExePath = Join-Path "outputs" $OutputExeName
Copy-Item -LiteralPath $LauncherExe -Destination $OutputExePath -Force

Write-Host "Built $OutputExePath"

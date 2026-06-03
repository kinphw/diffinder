# Diffinder build script
# Imports the VBA modules + custom ribbon into a fresh workbook and saves it as
# build\Diffinder.xlam.
#
# Requirements:
#   - Microsoft Excel installed (uses COM automation)
#   - Excel Trust Center: "Trust access to the VBA project object model" enabled
#     (HKCU\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 1)
#
# Usage:  pwsh -File build.ps1

$ErrorActionPreference = "Stop"
$root  = $PSScriptRoot
$out   = Join-Path $root "build\Diffinder.xlam"
$ui    = Join-Path $root "ribbon\customUI14.xml"
$stage = Join-Path $root "build\_src"

Add-Type -AssemblyName System.IO.Compression.FileSystem

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Excel is running. Close it (it may hold the add-in open) and retry."
}

New-Item -ItemType Directory -Force (Split-Path $out) | Out-Null
if (Test-Path $out) { Remove-Item $out -Force }

# --- SSOT: read version metadata and stage sources with injected values -------
$meta = Get-Content (Join-Path $root "version.json") -Raw | ConvertFrom-Json
Write-Host ("Version: {0} {1} (author: {2})" -f $meta.name, $meta.version, $meta.author)

if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force $stage | Out-Null

# ModDiff: copy bytes as-is (CP949 Korean — must not be re-encoded)
Copy-Item (Join-Path $root "src\ModDiff.bas") (Join-Path $stage "ModDiff.bas") -Force

# ModRibbon: inject SSOT values into the Public Const literals (ASCII-only file)
$rib = Get-Content (Join-Path $root "src\ModRibbon.bas") -Raw
$rib = $rib -replace '(?m)^(Public Const APP_NAME As String = ").*(")$',    ('${1}' + $meta.name    + '${2}')
$rib = $rib -replace '(?m)^(Public Const APP_VERSION As String = ").*(")$', ('${1}' + $meta.version + '${2}')
$rib = $rib -replace '(?m)^(Public Const APP_AUTHOR As String = ").*(")$',  ('${1}' + $meta.author  + '${2}')
[System.IO.File]::WriteAllText((Join-Path $stage "ModRibbon.bas"), $rib, (New-Object System.Text.ASCIIEncoding))

# 1) Import staged modules and save as .xlam (FileFormat 55 = xlOpenXMLAddIn)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    $wb.VBProject.VBComponents.Import((Join-Path $stage "ModDiff.bas"))   | Out-Null
    $wb.VBProject.VBComponents.Import((Join-Path $stage "ModRibbon.bas")) | Out-Null
    $wb.SaveAs($out, 55)
    $wb.Close($false)
} finally {
    $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
}

# 2) Inject the custom ribbon part + relationship into the OOXML package
$uiBytes = [System.IO.File]::ReadAllBytes($ui)
$zip = [System.IO.Compression.ZipFile]::Open($out, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    ($zip.Entries | Where-Object { $_.FullName -eq "customUI/customUI14.xml" }) | ForEach-Object { $_.Delete() }
    $e = $zip.CreateEntry("customUI/customUI14.xml")
    $s = $e.Open(); $s.Write($uiBytes, 0, $uiBytes.Length); $s.Close()

    $rels = $zip.Entries | Where-Object { $_.FullName -eq "_rels/.rels" } | Select-Object -First 1
    $rd = New-Object System.IO.StreamReader($rels.Open()); $relsXml = $rd.ReadToEnd(); $rd.Close()
    if ($relsXml -notmatch "customUI/customUI14.xml") {
        $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        $newRel  = '<Relationship Id="rIdDiffUI" Type="' + $relType + '" Target="customUI/customUI14.xml"/>'
        $relsXml = $relsXml -replace '</Relationships>', ($newRel + '</Relationships>')
        $rels.Delete()
        $r2 = $zip.CreateEntry("_rels/.rels")
        $sw = New-Object System.IO.StreamWriter($r2.Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write($relsXml); $sw.Close()
    }
} finally { $zip.Dispose() }

Write-Host ("Built: " + $out + "  (" + (Get-Item $out).Length + " bytes)")

# 3) Package a distributable zip: dist\diffinder_v<version>.zip (xlam + README)
$dist = Join-Path $root "dist"
New-Item -ItemType Directory -Force $dist | Out-Null
$zip = Join-Path $dist ("{0}_v{1}.zip" -f $meta.name.ToLower(), $meta.version)
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $out, (Join-Path $root "README.md") -DestinationPath $zip -Force
Write-Host ("Packaged: " + $zip + "  (" + (Get-Item $zip).Length + " bytes)")

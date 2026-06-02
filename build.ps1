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
$root = $PSScriptRoot
$out  = Join-Path $root "build\Diffinder.xlam"
$ui   = Join-Path $root "ribbon\customUI14.xml"

Add-Type -AssemblyName System.IO.Compression.FileSystem

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Excel is running. Close it (it may hold the add-in open) and retry."
}

New-Item -ItemType Directory -Force (Split-Path $out) | Out-Null
if (Test-Path $out) { Remove-Item $out -Force }

# 1) Import modules and save as .xlam (FileFormat 55 = xlOpenXMLAddIn)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    $wb.VBProject.VBComponents.Import((Join-Path $root "src\ModDiff.bas"))   | Out-Null
    $wb.VBProject.VBComponents.Import((Join-Path $root "src\ModRibbon.bas")) | Out-Null
    $wb.SaveAs($out, 55)
    $wb.Close($false)
} finally {
    $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
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

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDir = ".",

    [Parameter(Mandatory = $false)]
    [string]$WorkerScript = ".\split_docx_final2.ps1",

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "",

    [Parameter(Mandatory = $false)]
    [switch]$Recurse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$TargetDir = [System.IO.Path]::GetFullPath($TargetDir)
$WorkerScript = [System.IO.Path]::GetFullPath($WorkerScript)

if (-not (Test-Path -LiteralPath $TargetDir)) {
    throw "TargetDir not found: $TargetDir"
}

if (-not (Test-Path -LiteralPath $WorkerScript)) {
    throw "WorkerScript not found: $WorkerScript"
}

if ($OutputRoot -eq "") {
    $OutputRoot = Join-Path $TargetDir "bulk_output"
}
$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null

Write-Host ("TargetDir   : " + $TargetDir) -ForegroundColor Cyan
Write-Host ("WorkerScript: " + $WorkerScript) -ForegroundColor Cyan
Write-Host ("OutputRoot  : " + $OutputRoot) -ForegroundColor Cyan
Write-Host ("Recurse     : " + $Recurse.IsPresent) -ForegroundColor Cyan
Write-Host ""

$gciParams = @{
    LiteralPath = $TargetDir
    Filter      = "*.docx"
    File        = $true
}
if ($Recurse) {
    $gciParams["Recurse"] = $true
}

$files = Get-ChildItem @gciParams | Sort-Object FullName

if ($files.Count -eq 0) {
    Write-Host "No .docx files found." -ForegroundColor Yellow
    exit 0
}

Write-Host ("Found " + $files.Count + " docx file(s).") -ForegroundColor Green
Write-Host ""

$ok = 0
$ng = 0
$results = @()

for ($i = 0; $i -lt $files.Count; $i++) {
    $file = $files[$i]
    $inputFile = $file.FullName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
    $outDir = Join-Path $OutputRoot $baseName

    Write-Host ("[{0}/{1}] START  {2}" -f ($i + 1), $files.Count, $inputFile) -ForegroundColor Magenta
    Write-Host ("          OUT -> " + $outDir) -ForegroundColor DarkGray

    New-Item -ItemType Directory -Path $outDir -Force | Out-Null

    try {
        & $WorkerScript -InputFile $inputFile -OutputDir $outDir

        Write-Host ("[{0}/{1}] OK     {2}" -f ($i + 1), $files.Count, $inputFile) -ForegroundColor Green
        $ok++

        $results += [pscustomobject]@{
            File   = $inputFile
            Output = $outDir
            Status = "OK"
        }
    }
    catch {
        Write-Host ("[{0}/{1}] FAIL   {2}" -f ($i + 1), $files.Count, $inputFile) -ForegroundColor Red
        Write-Host ("          " + $_.Exception.Message) -ForegroundColor Yellow
        $ng++

        $results += [pscustomobject]@{
            File   = $inputFile
            Output = $outDir
            Status = "FAIL"
        }
    }

    Write-Host ""
}

Write-Host "===== SUMMARY =====" -ForegroundColor Cyan
Write-Host ("OK : " + $ok) -ForegroundColor Green
Write-Host ("NG : " + $ng) -ForegroundColor Red
Write-Host ("OUT: " + $OutputRoot) -ForegroundColor Cyan
Write-Host ""

$results | Format-Table -AutoSize
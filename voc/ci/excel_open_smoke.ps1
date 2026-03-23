param (
    [Parameter(Mandatory=$false)]
    [string]$FilePath,

    [Parameter(Mandatory=$false)]
    [string]$SearchRoot = ".",

    [Parameter(Mandatory=$false)]
    [string]$ArtifactPattern = "*.xlsx"
)

$ErrorActionPreference = "Stop"

function Write-Result {
    param([bool]$Success, [string]$Message, [object]$Details = @{})
    $evidencePath = ".sisyphus/evidence/task-8-excel-smoke.json"
    $evidenceDir = Split-Path -Parent $evidencePath
    if ($evidenceDir -and -not (Test-Path $evidenceDir)) {
        New-Item -ItemType Directory -Path $evidenceDir -Force | Out-Null
    }
    $result = @{
        timestamp = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
        success = $Success
        message = $Message
        repair_detected = ($Details.RepairDetected -eq $true)
        details = $Details
    }
    $result | ConvertTo-Json -Depth 10 | Out-File -FilePath $evidencePath -Encoding utf8
    if (-not $Success) { exit 1 }
    exit 0
}

function Resolve-ArtifactPath {
    param(
        [string]$CandidatePath,
        [string]$Root,
        [string]$Pattern
    )

    if ($CandidatePath) {
        $resolvedCandidate = [System.IO.Path]::GetFullPath($CandidatePath)
        if (Test-Path $resolvedCandidate) {
            return $resolvedCandidate
        }
        throw "Artifact not found at explicit path: $resolvedCandidate"
    }

    $resolvedRoot = [System.IO.Path]::GetFullPath($Root)
    if (-not (Test-Path $resolvedRoot)) {
        throw "Search root not found: $resolvedRoot"
    }

    $matches = Get-ChildItem -Path $resolvedRoot -Recurse -File -Filter $Pattern |
        Sort-Object LastWriteTimeUtc -Descending
    if (-not $matches -or $matches.Count -eq 0) {
        throw "No artifact matched pattern '$Pattern' under root '$resolvedRoot'"
    }
    return $matches[0].FullName
}

function Get-RepairEvidenceSnapshot {
    $roots = @(
        "$env:TEMP",
        "$env:LOCALAPPDATA\Temp",
        "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\Content.MSO",
        "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\Content.Excel"
    ) | Where-Object { $_ -and (Test-Path $_) }

    $patterns = @(
        "*repair*.xml",
        "*repaired*.xml",
        "*error*.xml",
        "*repair*.log",
        "*repaired*.log"
    )

    $snapshot = @{}
    foreach ($root in $roots) {
        foreach ($pattern in $patterns) {
            $files = Get-ChildItem -Path $root -Recurse -File -Filter $pattern -ErrorAction SilentlyContinue
            foreach ($file in $files) {
                $snapshot[$file.FullName] = @{
                    LastWriteTimeUtc = $file.LastWriteTimeUtc
                    Length = $file.Length
                }
            }
        }
    }
    return $snapshot
}

function Get-NewRepairEvidenceFiles {
    param(
        [hashtable]$Before,
        [hashtable]$After,
        [datetime]$SinceUtc
    )

    $newFiles = @()
    foreach ($path in $After.Keys) {
        if (-not $Before.ContainsKey($path)) {
            if ($After[$path].LastWriteTimeUtc -ge $SinceUtc) {
                $newFiles += $path
            }
            continue
        }
        $beforeMeta = $Before[$path]
        $afterMeta = $After[$path]
        if ($afterMeta.LastWriteTimeUtc -gt $beforeMeta.LastWriteTimeUtc -or $afterMeta.Length -ne $beforeMeta.Length) {
            if ($afterMeta.LastWriteTimeUtc -ge $SinceUtc) {
                $newFiles += $path
            }
        }
    }
    return $newFiles | Sort-Object -Unique
}

try {
    $absPath = Resolve-ArtifactPath -CandidatePath $FilePath -Root $SearchRoot -Pattern $ArtifactPattern
} catch {
    Write-Result -Success $false -Message $_.Exception.Message
}

Write-Host "Resolved workbook artifact: $absPath"

$beforeSnapshot = Get-RepairEvidenceSnapshot
$startedAtUtc = (Get-Date).ToUniversalTime()
$excel = $null
$workbook = $null
$repairWorkbook = $null

try {
    Write-Host "Opening Excel COM Object..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $normalOpenSucceeded = $false
    $normalOpenError = $null
    $repairOpenSucceeded = $false
    $repairOpenError = $null

    try {
        Write-Host "Opening workbook in normal mode..."
        $workbook = $excel.Workbooks.Open($absPath, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true, [Type]::Missing, [Type]::Missing, $false, $false, [Type]::Missing, $false, $true, 1)
        $sheetCount = $workbook.Worksheets.Count
        $normalOpenSucceeded = $true
        Write-Host "Normal open succeeded. Worksheet count: $sheetCount"
        $workbook.Close($false)
    } catch {
        $normalOpenError = $_.Exception.ToString()
        Write-Host "Normal open failed. Will attempt explicit repair mode."
    }

    if (-not $normalOpenSucceeded) {
        try {
            $repairWorkbook = $excel.Workbooks.Open($absPath, 0, $true, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true, [Type]::Missing, [Type]::Missing, $false, $false, [Type]::Missing, $false, $true, 2)
            $repairSheetCount = $repairWorkbook.Worksheets.Count
            $repairOpenSucceeded = $true
            Write-Host "Repair-mode open succeeded. Worksheet count: $repairSheetCount"
            $repairWorkbook.Close($false)
        } catch {
            $repairOpenError = $_.Exception.ToString()
        }
    }

    $afterSnapshot = Get-RepairEvidenceSnapshot
    $newRepairEvidence = Get-NewRepairEvidenceFiles -Before $beforeSnapshot -After $afterSnapshot -SinceUtc $startedAtUtc

    $repairDetected = ($newRepairEvidence.Count -gt 0) -or ((-not $normalOpenSucceeded) -and $repairOpenSucceeded)
    if ($repairDetected) {
        Write-Result -Success $false -Message "Excel repair evidence detected." -Details @{
            RepairDetected = $true
            ResolvedArtifactPath = $absPath
            NormalOpenSucceeded = $normalOpenSucceeded
            NormalOpenError = $normalOpenError
            RepairOpenSucceeded = $repairOpenSucceeded
            RepairOpenError = $repairOpenError
            RepairEvidenceFiles = $newRepairEvidence
        }
    }

    if (-not $normalOpenSucceeded) {
        Write-Result -Success $false -Message "Workbook failed to open in normal Excel mode." -Details @{
            RepairDetected = $false
            ResolvedArtifactPath = $absPath
            NormalOpenSucceeded = $normalOpenSucceeded
            NormalOpenError = $normalOpenError
            RepairOpenSucceeded = $repairOpenSucceeded
            RepairOpenError = $repairOpenError
            RepairEvidenceFiles = $newRepairEvidence
        }
    }

    Write-Result -Success $true -Message "Workbook opened in normal Excel mode with no repair evidence." -Details @{
        RepairDetected = $false
        ResolvedArtifactPath = $absPath
        NormalOpenSucceeded = $normalOpenSucceeded
        RepairOpenSucceeded = $repairOpenSucceeded
        RepairEvidenceFiles = $newRepairEvidence
    }

} catch {
    Write-Result -Success $false -Message "Critical error during Excel COM execution: $($_.Exception.Message)" -Details @{ Error = $_.Exception.ToString() }
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $repairWorkbook) {
        try { $repairWorkbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }
}

#requires -Version 5.1

param(
    [string]$CommitNote = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepositoryPath = Split-Path -Parent $MyInvocation.MyCommand.Path

function Test-GitAvailable {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw 'git command not found in PATH.'
    }
}

function Get-StatusInfo {
    param(
        [string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line) -or $Line.Length -lt 4) {
        return $null
    }

    $indexStatus = $Line.Substring(0, 1)
    $worktreeStatus = $Line.Substring(1, 1)
    $pathPart = $Line.Substring(3).Trim()

    if ([string]::IsNullOrWhiteSpace($pathPart)) {
        return $null
    }

    return [pscustomobject]@{
        IndexStatus    = $indexStatus
        WorktreeStatus = $worktreeStatus
        PathPart       = $pathPart
    }
}

function Format-ListSummary {
    param(
        [string[]]$Items
    )

    if (-not $Items -or $Items.Count -eq 0) {
        return $null
    }

    if ($Items.Count -le 3) {
        return ($Items -join ', ')
    }

    $previewItems = $Items | Select-Object -First 3
    $preview = $previewItems -join ', '
    $remaining = $Items.Count - $previewItems.Count
    return ('{0} and {1} more item{2}' -f $preview, $remaining, $(if ($remaining -gt 1) { 's' } else { '' }))
}

function Get-ChangeSummary {
    param(
        [string[]]$StatusLines
    )

    $added = New-Object System.Collections.Generic.List[string]
    $modified = New-Object System.Collections.Generic.List[string]
    $deleted = New-Object System.Collections.Generic.List[string]
    $renamed = New-Object System.Collections.Generic.List[string]

    foreach ($line in $StatusLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $info = Get-StatusInfo -Line $line
        if (-not $info) {
            continue
        }

        $statuses = @($info.IndexStatus, $info.WorktreeStatus)

        if ($statuses -contains '?' ) {
            $added.Add($info.PathPart)
            continue
        }

        if ($statuses -contains 'R') {
            if ($info.PathPart -match '^(.*?)->(.*)$') {
                $from = $matches[1].Trim()
                $to = $matches[2].Trim()
                $renamed.Add(('{0} -> {1}' -f $from, $to))
            }
            else {
                $renamed.Add($info.PathPart)
            }
            continue
        }

        if ($statuses -contains 'A') {
            $added.Add($info.PathPart)
            continue
        }

        if ($statuses -contains 'D') {
            $deleted.Add($info.PathPart)
            continue
        }

        if ($statuses -contains 'M') {
            $modified.Add($info.PathPart)
            continue
        }

        $modified.Add($info.PathPart)
    }

    $segments = New-Object System.Collections.Generic.List[string]

    $addedSummary = Format-ListSummary -Items $added.ToArray()
    if ($addedSummary) { $segments.Add("Added: $addedSummary") }

    $modifiedSummary = Format-ListSummary -Items $modified.ToArray()
    if ($modifiedSummary) { $segments.Add("Modified: $modifiedSummary") }

    $deletedSummary = Format-ListSummary -Items $deleted.ToArray()
    if ($deletedSummary) { $segments.Add("Deleted: $deletedSummary") }

    $renamedSummary = Format-ListSummary -Items $renamed.ToArray()
    if ($renamedSummary) { $segments.Add("Renamed: $renamedSummary") }

    if ($segments.Count -eq 0) {
        return 'Updated project files'
    }

    $summary = $segments -join '; '
    if ($summary.Length -gt 80) {
        return ($summary.Substring(0, 77) + '...')
    }

    return $summary
}

try {
    Test-GitAvailable
    Set-Location -Path $RepositoryPath

    $statusResult = git status --short
    if ($statusResult -is [string]) {
        $statusLines = @($statusResult)
    }
    else {
        $statusLines = @()
        foreach ($line in $statusResult) {
            $statusLines += $line
        }
    }

    if (-not $statusLines -or $statusLines.Count -eq 0) {
        Write-Host 'No changes to commit.'
        exit 0
    }

    git add .

    $today = (Get-Date).ToString('yyyyMMdd')
    $existingSubjects = git log --format='%s' --grep "^V$today\."

    $maxIndex = 0
    foreach ($subject in $existingSubjects) {
        if ($subject -match "^V$today\.([0-9]+)_") {
            $number = [int]$matches[1]
            if ($number -gt $maxIndex) {
                $maxIndex = $number
            }
        }
    }

    $nextIndex = $maxIndex + 1
    $indexText = $nextIndex.ToString('00')

    $summary = Get-ChangeSummary -StatusLines $statusLines
    if ($CommitNote) {
        $summaryWithNote = "$summary ($CommitNote)"
        if ($summaryWithNote.Length -gt 80) {
            $summary = $summaryWithNote.Substring(0, 77) + '...'
        }
        else {
            $summary = $summaryWithNote
        }
    }
    $commitMessage = "V{0}.{1}_{2}" -f $today, $indexText, $summary

    git commit -m $commitMessage
    Write-Host ("Commit created: {0}" -f $commitMessage)
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}

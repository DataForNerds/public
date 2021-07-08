$msData = Invoke-WebRequest "https://winreleaseinfoprod.blob.core.windows.net/winreleaseinfoprod/en-US.html" -UseBasicParsing
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/mswin/releases.json" | Select -ExpandProperty Content | ConvertFrom-Json

$allReleases = [RegEx]::New('(?msi)<span(?:[^>])class="triangle".*?>&#9660;<\/span>(?: |)<strong>Version (.*?) \(OS Build (\d{1,})\)<\/strong>').Matches($msData.RawContent)

$releaseList = New-Object System.Collections.ArrayList

$allReleases.ForEach{
    $releaseList.add(
        [PSCustomObject]@{
            Version = $_.Groups[2].value
            FullVersion = "10.0.$($_.Groups[2].Value)"
            Build = $_.Groups[1].Value
        }
    ) | Out-Null
}

$releaseList = $releaseList | Sort-Object Version -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @("https://docs.microsoft.com/en-us/windows/release-health/release-information","https://winreleaseinfoprod.blob.core.windows.net/winreleaseinfoprod/en-US.html")
    }
    "Data" = $releaseList
}

$diffSinceLastUpdate = New-TimeSpan -Start $d4nData.DataForNerds.LastUpdatedUTC -End (Get-Date).ToUniversalTime()
$allProperties = $releaseList[0].psobject.Properties.Name

If($diffSinceLastUpdate.TotalDays -ge 7 -or (Compare-Object $d4nData.Data $releaseList -Property $allProperties -SyncWindow 0)) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
    $outputFile = Join-Path $outputFolder -ChildPath "releases.json"

    $jsonData = $outputData | ConvertTo-Json -Compress 
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)   
} else {
    Write-Host "The data has not changed and it's only been $([math]::Round($diffSinceLastUpdate.TotalDays,2)) day(s) since the last update."
}

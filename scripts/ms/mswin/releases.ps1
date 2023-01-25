$msdata = @("https://learn.microsoft.com/en-us/windows/release-health/release-information",
            "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information")
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/mswin/releases.json" | Select-Object -ExpandProperty Content | ConvertFrom-Json

$releaseList = New-Object System.Collections.ArrayList

foreach ($sourceURL in $msdata) {
    $source = Invoke-WebRequest $sourceURL -UseBasicParsing
    $allReleases = [RegEx]::New('(?msi)<strong>Version (.*?) \(OS Build (\d{1,})\)<\/strong>').Matches($source.RawContent)
    $allReleases.ForEach{
        $releaseList.add(
            [PSCustomObject]@{
                Version = $_.Groups[2].value
                FullVersion = "10.0.$($_.Groups[2].Value)"
                Build = $_.Groups[1].Value
            }
        ) | Out-Null
    }
}

$releaseList = $releaseList | Sort-Object Version | Select-Object Version,FullVersion,Build -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @("https://docs.microsoft.com/en-us/windows/release-health/release-information","https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information")
    }
    "Data" = $releaseList
}

$allProperties = $releaseList[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $releaseList -Property $allProperties -SyncWindow 0) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
    $outputFile = Join-Path $outputFolder -ChildPath "releases.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)   
} else {
    Write-Host "The data has not changed."
}

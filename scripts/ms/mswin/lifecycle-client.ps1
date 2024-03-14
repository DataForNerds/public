$msdata = @("https://learn.microsoft.com/en-us/lifecycle/products/windows-11-enterprise-and-education")
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/mswin/lifecycle-client.json" | Select-Object -ExpandProperty Content | ConvertFrom-Json

$releaseList = New-Object System.Collections.ArrayList

foreach ($sourceURL in $msdata) {
    $source = Invoke-WebRequest $sourceURL -UseBasicParsing
    $allReleases = [RegEx]::New(('<td>Version (.*?)<\/td>\s*<td align="right">\s*<local-time timezone="America\/Los_Angeles" format="date" datetime="(.*?)">(.*?)<\/local-time>\s*<\/td>\s*<td align="right">\s*<local-time timezone="America\/Los_Angeles" format="date" datetime="(.*?)">(.*?)<\/local-time>')).Matches($source.RawContent)
    $allReleases.ForEach{
        $releaseList.add(
            [PSCustomObject]@{
                Version = $_.Groups[1].value
                StartDate = $(Get-Date $_.Groups[2].Value -Format "yyyy-MM-dd")
                EndDate = $(Get-Date $_.Groups[3].Value -Format "yyyy-MM-dd")
            }
        ) | Out-Null
    }
}

$releaseList = $releaseList | Sort-Object Version | Select-Object Version,StartDate,EndDate -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @("https://learn.microsoft.com/en-us/lifecycle/products/windows-11-enterprise-and-education")
    }
    "Data" = $releaseList
}

$allProperties = $releaseList[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $releaseList -Property $allProperties -SyncWindow 0) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
    $outputFile = Join-Path $outputFolder -ChildPath "lifecycle-client.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)   
} else {
    Write-Host "The data has not changed."
}

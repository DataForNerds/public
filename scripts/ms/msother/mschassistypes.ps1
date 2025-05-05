$msdata = @("https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/cim-chassis")
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/msother/mschassistypes.json" | Select-Object -ExpandProperty Content | ConvertFrom-Json

$releaseList = New-Object System.Collections.ArrayList

foreach ($sourceURL in $msdata) {
    $source = Invoke-WebRequest $sourceURL -UseBasicParsing
    $focusedSource = $source.RawContent.ToString().Split("<p><strong>CreationClassName</strong></p>")
    $allReleases = [RegEx]::New(('\<strong\>(.*?)\<\/strong\>\s\((.*?)\)')).Matches($focusedSource[0])
    $allReleases.ForEach{
        $releaseList.add(
            [PSCustomObject]@{
                Value = $_.Groups[1].Value
                Number = $_.Groups[2].Value
            }
        ) | Out-Null
    }
}

$releaseList = $releaseList | Sort-Object Number | Select-Object Value,Number -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @("https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/cim-chassis")
    }
    "Data" = $releaseList
}

$allProperties = $releaseList[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $releaseList -Property $allProperties -SyncWindow 0) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/msother")
    $outputFile = Join-Path $outputFolder -ChildPath "mschassistypes.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)   
} else {
    Write-Host "The data has not changed."
}

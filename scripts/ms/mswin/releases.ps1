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

## Fix issue where 19044 (21H2) still now showing up in the MS Blob
if($releaseList.Version -notcontains "19044") {
    $releaseList += [PSCustomObject]@{
        Version = "19044"
        FullVersion = "10.0.19044"
        Build = "21H2"
    }
}

$releaseList = $releaseList | Sort-Object Version | Select-Object Version,FullVersion,Build -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @("https://docs.microsoft.com/en-us/windows/release-health/release-information","https://winreleaseinfoprod.blob.core.windows.net/winreleaseinfoprod/en-US.html")
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

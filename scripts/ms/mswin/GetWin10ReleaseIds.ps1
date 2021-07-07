$msData = Invoke-WebRequest "https://winreleaseinfoprod.blob.core.windows.net/winreleaseinfoprod/en-US.html" -UseBasicParsing

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

$outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
$releaseList | ConvertTo-Json | Out-File (Join-Path $outputFolder -ChildPath "releases.json") -Encoding utf8
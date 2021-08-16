$rootPage = "https://docs.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date"
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/msapps/buildnumbers.json" | Select -ExpandProperty Content | ConvertFrom-Json

$pageData = Invoke-WebRequest $rootPage -UseBasicParsing

If($pageData.StatusCode -ne 200) {
    Throw "Error $($pageData.StatusCode) Getting Page Data"
}

$tables = [regex]::New('(?msi)<table>(?:.*?)<tbody>(.*?)<\/tbody>').Matches($pageData.Content)

$versionHistoryRows = [regex]::New('(?msi)<tr>(.*?)<\/tr>').Matches($tables[1].Groups[1].Value)

$m365Releases = New-Object System.Collections.ArrayList

$rxInnerLink = [Regex]::New('(?msi)<a(?:[^>])*>(.*?)<\/a>')
$rxVersionBuild = [Regex]::New('(?msi)Version (.*?) \(Build {1,}(.*?)\)')

$versionHistoryRows.ForEach{

    $cellData = [regex]::New('(?msi)<td(?:[^>]*)>(.*?)<\/td>').Matches($_.Groups[1].Value)
    $releaseYear = $cellData[0].Groups[1].Value
    $releaseDate = $($cellData[1].Groups[1].Value -replace "<br(?:[^>])*>","").Trim()

    if(-Not($releaseYear)) {
        $releaseYear = $lastReleaseYear
    } else {
        $lastReleaseYear = $releaseYear
    }

    $release = $(Get-Date "$releaseDate $releaseYear" -Format "yyyy-MM-dd")

    $channelCurrentLinks = $rxInnerLink.Matches($cellData[2].Groups[1].Value)
    $channelMonthlyEnterpriseLinks = $rxInnerLink.Matches($cellData[3].Groups[1].Value)
    $channelSACEnterprisePreviewLinks = $rxInnerLink.Matches($cellData[4].Groups[1].Value)
    $channelSACEnterpriseLinks = $rxInnerLink.Matches($cellData[5].Groups[1].Value)

    $allLinks = New-Object System.Collections.ArrayList

    $allLinks.AddRange(@($channelCurrentLinks.groups.where{$_.Name -eq 1} | Select-Object @{Name="Channel";Expression={"Current"}},@{Name="Value";Expression={$_.Value}}))
    $allLinks.AddRange(@($channelMonthlyEnterpriseLinks.groups.where{$_.Name -eq 1} | Select-Object @{Name="Channel";Expression={"Monthly Enterprise"}},@{Name="Value";Expression={$_.Value}}))
    $allLinks.AddRange(@($channelSACEnterprisePreviewLinks.groups.where{$_.Name -eq 1} | Select-Object @{Name="Channel";Expression={"Semi-Annual Enterprise Preview"}},@{Name="Value";Expression={$_.Value}}))
    $allLinks.AddRange(@($channelSACEnterpriseLinks.groups.where{$_.Name -eq 1} | Select-Object @{Name="Channel";Expression={"Semi-Annual Enterprise"}},@{Name="Value";Expression={$_.Value}}))

    $allLinks.ForEach{
        $versionBuild = $rxVersionBuild.Matches($_.Value)
            
        $thisChannel = $_.Channel

        $versionBuild.ForEach{
            $m365Releases.Add(
                [PSCustomObject]@{
                    ReleaseDate = $release
                    Channel = $thisChannel
                    Build = $_.Groups[2].Value
                    Version = $_.Groups[1].Value
                    FullBuild = "16.0.$($_.Groups[2].Value)"
                }
            ) | Out-Null
        }
    }

}

$m365Releases = $m365Releases | Sort-Object ReleaseDate -Descending | Select-Object ReleaseDate,Channel,Build,Version,FullBuild -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @($rootPage)
    }
    "Data" = $m365Releases
}

$allProperties = $m365Releases[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $outputData.Data -Property $allProperties) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/msapps")
    $outputFile = Join-Path $outputFolder -ChildPath "buildnumbers.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)
} else {
    Write-Host "The data has not changed."
}


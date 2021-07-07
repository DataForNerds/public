$rootPage = "https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11"

$pageData = Invoke-WebRequest $rootPage -UseBasicParsing

If($pageData.StatusCode -ne 200) {
    Throw "Erorr $($pageData.StatusCode) Getting Page Data"
}

$pageContent = $pageData.Content

$rxBuildata = [regex]::New('(?i)>(.*?) \(OS Build.*?(\d.*?)\)(.*?)<')
$rxPatchDesc = [regex]::New('(?i)^(.*? \d{1,2}[,|] \d{4}).*?(K.*)$')

$buildData = $rxBuildata.Matches($pageContent)

$patchList = New-Object System.Collections.ArrayList

$buildData.ForEach{
    
    $versionList = $_.Groups[2].Value
    $versionList = $versionList -replace "and",","
    $versionList = $versionList -replace " ",""
    $versionList = $versionList -split "," | Where-Object { $_ -ne "" }
    
    $description = [System.Web.HttpUtility]::HtmlDecode($_.Groups[1].Value)

    $PatchInfo = $rxPatchDesc.Matches($description)

    ForEach($version in $versionList) {
        $patchList.add(
            [PSCustomObject]@{
                # "Match"=$_.Groups[0].Value
                # "Description"=$description
                "Win10Version"="10.0.$version"
                "Version"=$version
                "ReleaseDate" = $(Get-Date $PatchInfo.groups[1].Value -Format "yyyy-MM-dd")
                "Article" = $PatchInfo.Groups[2].Value
                "Comment"=$_.Groups[3].Value.ToString().Trim()
                "Source"=$rootPage
            }
        ) | Out-Null
    }
}

$patchList = $patchList | Sort-Object Win10Version -Unique

$outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
$patchList | ConvertTo-Json | Out-File (Join-Path $outputFolder -ChildPath "buildnumbers.json")


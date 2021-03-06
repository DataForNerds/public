$rootPage = "https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11"
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/mswin/buildnumbers.json" | Select-Object -ExpandProperty Content | ConvertFrom-Json

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
    
    $KBTitle = $_.Groups[0].Value.SubString(1)
    $KBTitle = $KBTitle.Substring(0,$KBTitle.Length-1)
    $KBTitle = [System.Web.HttpUtility]::HtmlDecode($KBTitle)

    $description = [System.Web.HttpUtility]::HtmlDecode($_.Groups[1].Value)

    $PatchInfo = $rxPatchDesc.Matches($description)

    $ArticleNumber = $PatchInfo.Groups[2].Value
    
    If($ArticleNumber -like "KB *") {
        # Fixes issue where there is sometimes a space after KB
        $ArticleNumber = "KB$($ArticleNumber.Substring(3))"
    }

    If($ArticleNumber -like "* *") {
        $ArticleNumber = $ArticleNumber -split " " | Select -First 1
    }

    ForEach($version in $versionList) {
        $patchList.add(
            [PSCustomObject]@{
                # "Match"=$_.Groups[0].Value
                # "Description"=$description
                "Win10Version"="10.0.$version"
                "Version"=$version
                "ReleaseDate" = $(Get-Date $PatchInfo.groups[1].Value -Format "yyyy-MM-dd")
                "Article" = $ArticleNumber
                "KBTitle" = $KBTitle
                "Comment"=$_.Groups[3].Value.ToString().Trim()
            }
        ) | Out-Null
    }
}

$patchList = $patchList | Sort-Object ReleaseDate | Select-Object Win10Version,Version,ReleaseDate,Article,KBTitle,Comment -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = "https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11"
    }
    "Data" = $patchList
}

$allProperties = $patchList[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $outputData.Data -Property $allProperties -SyncWindow 0) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
    $outputFile = Join-Path $outputFolder -ChildPath "buildnumbers.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)
} else {
    Write-Host "The data has not changed."
}



# 808
$patchList | FT -AutoSize


$patchList | Where-Object { $_.Article -like "*4580346*" }

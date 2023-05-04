$rootPages = @("https://aka.ms/WindowsUpdateHistory",
               "https://aka.ms/Windows11UpdateHistory",
               "https://support.microsoft.com/en-us/topic/windows-server-2022-update-history-e1caa597-00c5-4ab9-9f3e-8212fe80b2ee")

$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/mswin/buildnumbers.json" | Select-Object -ExpandProperty Content | ConvertFrom-Json

$patchList = New-Object System.Collections.ArrayList

foreach ($sourceURL in $rootPages) {
    $pageData = Invoke-WebRequest $sourceURL -UseBasicParsing

    if($pageData.StatusCode -ne 200) {
        Throw "Error $($pageData.StatusCode) Getting Page Data"
    }
    
    $pageContent = $pageData.Content
    
    $rxBuildata = [regex]::New('(?i)>(.*?) \(OS Build.*?(\d.*?)\)(.*?)<')
    $rxPatchDesc = [regex]::New('(?i)^(.*? \d{1,2}[,|] \d{4}).*?(K.*)$')
    
    $buildData = $rxBuildata.Matches($pageContent)

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
            $ArticleNumber = $ArticleNumber -split " " | Select-Object -First 1
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
                    "LTSCOnly" = $False
                    "Comment"=$_.Groups[3].Value.ToString().Trim()
                }
            ) | Out-Null
        }
    }

}

$patchList = $patchList | Sort-Object ReleaseDate | Select-Object Win10Version,Version,ReleaseDate,Article,KBTitle,LTSCOnly,Comment -Unique

# Manual Updates to List
$patchList = $patchList | Where-Object { $_.Win10Version -notin ('10.0.14393.5127') }  # Windows Server Only
$patchList | Where-Object { $_.Win10Version -like "10.0.17763.*" -and (Get-Date $_.ReleaseDate) -ge (Get-Date '2021-05-12') } | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name LTSCOnly -Value $True -Force }


$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = "https://aka.ms/WindowsUpdateHistory","https://aka.ms/Windows11UpdateHistory","https://support.microsoft.com/en-us/topic/windows-server-2022-update-history-e1caa597-00c5-4ab9-9f3e-8212fe80b2ee"
    }
    "Data" = $patchList
}

$allProperties = $patchList[0].psobject.Properties.Name

if(Compare-Object $d4nData.Data $outputData.Data -Property $allProperties -SyncWindow 0) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/mswin")
    $outputFile = Join-Path $outputFolder -ChildPath "buildnumbers.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)
} else {
    Write-Host "The data has not changed."
}
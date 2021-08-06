$rootPage = "https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a"
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/msother/mslocales.json" | Select -ExpandProperty Content | ConvertFrom-Json

$pageData = Invoke-WebRequest $rootPage -UseBasicParsing

If($pageData.StatusCode -ne 200) {
    Throw "Error $($pageData.StatusCode) Getting Page Data"
}

$tables = [regex]::New('(?msi)<table>(?:.*?)<tbody>(.*?)<\/tbody>').Matches($pageData.Content)

$localeRows = [regex]::New('(?msi)<tr>(.*?)<\/tr>').Matches($tables[0].Groups[1].Value)

$locales = New-Object System.Collections.ArrayList

$localeRows.ForEach{
    
    $cellData = [regex]::New('(?msi)<td(?:[^>]*)>(.*?)<\/td>').Matches($_.Groups[1].Value)
    
    $LangCode = $cellData[0].Groups[1].value.Trim() -replace "(<(?:.*?)>)",""
    $LangName = $cellData[1].Groups[1].Value.Trim() -replace "(<(?:.*?)>)",""
    $LangTag = $cellData[2].Groups[1].Value.Trim() -replace "(<(?:.*?)>)",""
    
    If($LangCode -ne "Any other value") {
        $locales.Add(
            [PSCustomObject]@{
                LangCode = [int]$LangCode
                LangName = $LangName
                LangTag = $LangTag
            }
        ) | Out-Null
    }

}

$locales = $locales | Sort-Object LangCode -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @($rootPage)
    }
    "Data" = $locales
}

$allProperties = $locales[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $outputData.Data -Property $allProperties) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/msother")
    $outputFile = Join-Path $outputFolder -ChildPath "mslocales.json"

    $jsonData = $outputData | ConvertTo-Json -Compress 
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)
} else {
    Write-Host "The data has not changed."
}
